using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Geometry3d;

namespace TeklaGroupFinder
{
    public partial class GroupFinderForm : Form
    {
        private Model MyModel;

        public GroupFinderForm()
        {
            InitializeComponent();
            MyModel = new Model();
            if (MyModel.GetConnectionStatus())
            {
                connectionStatusPanel.BackColor = System.Drawing.Color.Green;
                statusLabel.Text = "Connected to: " + MyModel.GetInfo().ModelName;
            }
            else
            {
                connectionStatusPanel.BackColor = System.Drawing.Color.Red;
                statusLabel.Text = "Not connected. Please open a Tekla model.";
            }
        }

        private void createSignatureButton_Click(object sender, EventArgs e)
        {
            if (MyModel == null || !MyModel.GetConnectionStatus()) return;

            var picker = new Picker();
            var selected = picker.PickObjects(Picker.PickObjectsEnum.PICK_N_PARTS, "Select Post and Cap Plate");
            var parts = new List<Part>();
            while (selected.MoveNext()) if (selected.Current is Part p) parts.Add(p);

            Part mainPost = parts.FirstOrDefault(p => !(p is ContourPlate));
            ContourPlate capPlate = parts.OfType<ContourPlate>().FirstOrDefault();

            if (mainPost == null || capPlate == null)
            {
                MessageBox.Show("Selection must include a post and a ContourPlate.", "Selection Error");
                return;
            }

            var groupName = Prompt("New Signature", "Enter Signature Name:", "Type_1");
            if (string.IsNullOrEmpty(groupName)) return;

            double.TryParse(toleranceTextBox.Text, out double distTolerance);

            GetPostEndpoints(mainPost, out Point startPt, out Point endPt);
            Point capCog = GetPartCog(capPlate);
            Point refNode = Distance.PointToPoint(startPt, capCog) <= Distance.PointToPoint(endPt, capCog) ? startPt : endPt;

            double volume = 0;
            capPlate.GetReportProperty("VOLUME", ref volume);

            var signature = new GroupSignature
            {
                GroupName = groupName,
                PostProfile = mainPost.Profile.ProfileString,
                CapPlateProfile = capPlate.Profile.ProfileString,
                PostName = mainPost.Name,
                CapPlateName = capPlate.Name,
                CornerDistances = GetCornerDistances(capPlate, refNode),
                CapPlateVolume = volume,
                PartCount = parts.Count,
                DistanceTolerance = distTolerance > 0 ? distTolerance : 2.0
            };

            var path = GetDataPath();
            Directory.CreateDirectory(path);
            File.WriteAllText(Path.Combine(path, $"{groupName}.json"), JsonConvert.SerializeObject(signature, Formatting.Indented));

            statusLabel.Text = $"Signature '{groupName}' Created.";
        }

        private const double ProximityLimitMm = 200.0;
        private const double VolumeToleranceMm3 = 2000.0;
        private const int ProgressUpdateInterval = 25;

        private class PlateCache
        {
            public ContourPlate Plate { get; set; }
            public Point Cog { get; set; }
            public string Profile { get; set; }
            public string Name { get; set; }
            public double Volume { get; set; }
        }

        private class PostCache
        {
            public Part Post { get; set; }
            public Point StartPoint { get; set; }
            public Point EndPoint { get; set; }
            public string Profile { get; set; }
            public string Name { get; set; }
        }

        private async void findMatchesButton_Click(object sender, EventArgs e)
        {
            if (MyModel == null || !MyModel.GetConnectionStatus()) return;
            var path = GetDataPath();

            if (!Directory.Exists(path)) return;

            var signatures = Directory.GetFiles(path, "*.json")
                .Select(f => JsonConvert.DeserializeObject<GroupSignature>(File.ReadAllText(f))).ToList();

            var validPostProfiles = new HashSet<string>(signatures.Where(s => !string.IsNullOrEmpty(s.PostProfile)).Select(s => s.PostProfile.Trim().ToUpper()));
            var validPlateProfiles = new HashSet<string>(signatures.Where(s => !string.IsNullOrEmpty(s.CapPlateProfile)).Select(s => s.CapPlateProfile.Trim().ToUpper()));
            var validPostNames = new HashSet<string>(signatures.Where(s => !string.IsNullOrEmpty(s.PostName)).Select(s => s.PostName.Trim().ToUpper()));
            var validPlateNames = new HashSet<string>(signatures.Where(s => !string.IsNullOrEmpty(s.CapPlateName)).Select(s => s.CapPlateName.Trim().ToUpper()));

            var targetPosts = new List<PostCache>();
            var rawPlates = new List<ContourPlate>();
            var iter = MyModel.GetModelObjectSelector().GetAllObjectsWithType(new Type[] { typeof(Part) });
            while (iter.MoveNext())
            {
                if (iter.Current is ContourPlate cp)
                {
                    string profileKey = cp.Profile.ProfileString.Trim().ToUpper();
                    string nameKey = cp.Name?.Trim().ToUpper() ?? "";
                    if (validPlateProfiles.Contains(profileKey) && (validPlateNames.Count == 0 || validPlateNames.Contains(nameKey)))
                        rawPlates.Add(cp);
                }
                else if (iter.Current is Part part)
                {
                    string profileKey = part.Profile.ProfileString.Trim().ToUpper();
                    string nameKey = part.Name?.Trim().ToUpper() ?? "";
                    if (validPostProfiles.Contains(profileKey) && (validPostNames.Count == 0 || validPostNames.Contains(nameKey)))
                    {
                        GetPostEndpoints(part, out Point startPt, out Point endPt);
                        targetPosts.Add(new PostCache { Post = part, StartPoint = startPt, EndPoint = endPt, Profile = part.Profile.ProfileString, Name = part.Name });
                    }
                }
            }

            var plateCaches = rawPlates.Select(cp =>
            {
                double vol = 0;
                cp.GetReportProperty("VOLUME", ref vol);
                return new PlateCache { Plate = cp, Cog = GetPartCog(cp), Profile = cp.Profile.ProfileString, Name = cp.Name, Volume = vol };
            }).ToList();

            processingProgressBar.Visible = true;
            processingProgressBar.Maximum = targetPosts.Count;
            int matchCount = 0;
            double.TryParse(toleranceTextBox.Text, out double globalTol);
            double proxLimitSq = ProximityLimitMm * ProximityLimitMm;

            await System.Threading.Tasks.Task.Run(() =>
            {
                for (int i = 0; i < targetPosts.Count; i++)
                {
                    if (i % ProgressUpdateInterval == 0 || i == targetPosts.Count - 1)
                        this.Invoke((MethodInvoker)delegate { processingProgressBar.Value = i + 1; });

                    PostCache postCache = targetPosts[i];
                    Part main = postCache.Post;
                    Point startPt = postCache.StartPoint;
                    Point endPt = postCache.EndPoint;

                    PlateCache bestPlate = null;
                    Point postNode = null;
                    double minDistSq = double.MaxValue;
                    foreach (PlateCache pc in plateCaches)
                    {
                        double distStartSq = DistanceSquared(pc.Cog, startPt);
                        double distEndSq = DistanceSquared(pc.Cog, endPt);
                        double distSq = Math.Min(distStartSq, distEndSq);
                        if (distSq < proxLimitSq && distSq < minDistSq)
                        {
                            minDistSq = distSq;
                            bestPlate = pc;
                            postNode = distStartSq <= distEndSq ? startPt : endPt;
                        }
                    }

                    if (bestPlate == null) continue;

                    string postProf = postCache.Profile;
                    string postName = postCache.Name;
                    string plateProf = bestPlate.Profile;
                    string plateName = bestPlate.Name;
                    var corners = GetCornerDistances(bestPlate.Plate, postNode);
                    double plateVolume = bestPlate.Volume;

                    foreach (var sig in signatures)
                    {
                        double activeTol = sig.DistanceTolerance > 0 ? sig.DistanceTolerance : globalTol;

                        if (DoSignaturesMatch(sig, postName, postProf, plateName, plateProf, plateVolume, corners, activeTol))
                        {
                            main.SetUserProperty(udaNameTextBox.Text, sig.GroupName);
                            main.Modify();
                            bestPlate.Plate.SetUserProperty(udaNameTextBox.Text, sig.GroupName);
                            bestPlate.Plate.Modify();
                            matchCount++;
                            break;
                        }
                    }

                    // To create a new Assembly grouping the post and plate together (for future reference):
                    // Assembly newAssembly = new Assembly();
                    // newAssembly.SetMainPart(main);
                    // newAssembly.Add(plate);
                    // newAssembly.Insert();
                }
            });

            MyModel.CommitChanges("Posts Grouped.");
            processingProgressBar.Visible = false;
            statusLabel.Text = $"Tagged {matchCount} posts.";
            MessageBox.Show($"Search complete. {matchCount} matches found.", "Result");
        }

        private void openCatalogueButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The Catalogue will be built in a later step.", "Not Implemented");
        }

        #region Helper Methods
        private bool DoSignaturesMatch(GroupSignature s, string postN, string postP, string plateN, string plateP, double volume, List<double> c, double t)
        {
            if (s.PostProfile.Trim().ToUpper() != postP.Trim().ToUpper() ||
                s.CapPlateProfile.Trim().ToUpper() != plateP.Trim().ToUpper()) return false;

            if (!string.IsNullOrEmpty(s.PostName) && s.PostName.Trim().ToUpper() != (postN ?? "").Trim().ToUpper()) return false;
            if (!string.IsNullOrEmpty(s.CapPlateName) && s.CapPlateName.Trim().ToUpper() != (plateN ?? "").Trim().ToUpper()) return false;

            if (Math.Abs(s.CapPlateVolume - volume) > VolumeToleranceMm3) return false;

            if (s.CornerDistances.Count != c.Count) return false;

            for (int i = 0; i < c.Count; i++)
                if (Math.Abs(s.CornerDistances[i] - c[i]) > t) return false;

            return true;
        }

        private string GetDataPath() => Path.Combine("J:\\", projectNumberTextBox.Text, "600 QA and QC", "608 Internal Documents", "MegaPanelPreFabGroup");

        private string GetProperty(Part p, string n) { string v = ""; p.GetReportProperty(n, ref v); return v; }

        private Point GetReportPoint(Part p, string prefix)
        {
            double x = 0, y = 0, z = 0;
            p.GetReportProperty(prefix + "_X", ref x);
            p.GetReportProperty(prefix + "_Y", ref y);
            p.GetReportProperty(prefix + "_Z", ref z);
            return new Point(x, y, z);
        }

        private void GetPostEndpoints(Part post, out Point start, out Point end)
        {
            if (post is Beam beam) { start = beam.StartPoint; end = beam.EndPoint; }
            else { start = GetReportPoint(post, "START"); end = GetReportPoint(post, "END"); }
        }

        private List<double> GetCornerDistances(ContourPlate p, Point r)
        {
            var d = new List<double>();
            foreach (ContourPoint cp in p.Contour.ContourPoints)
                d.Add(Math.Round(Distance.PointToPoint(r, new Point(cp.X, cp.Y, cp.Z)), 1));
            d.Sort(); return d;
        }

        private static double DistanceSquared(Point p1, Point p2)
        {
            double dx = p1.X - p2.X;
            double dy = p1.Y - p2.Y;
            double dz = p1.Z - p2.Z;
            return (dx * dx) + (dy * dy) + (dz * dz);
        }

        private Point GetPartCog(ModelObject m)
        {
            if (m is Part part)
            {
                var solid = part.GetSolid();
                if (solid != null)
                {
                    return new Point(
                        (solid.MinimumPoint.X + solid.MaximumPoint.X) / 2.0,
                        (solid.MinimumPoint.Y + solid.MaximumPoint.Y) / 2.0,
                        (solid.MinimumPoint.Z + solid.MaximumPoint.Z) / 2.0
                    );
                }
            }

            // Fallback to report properties
            double x = 0, y = 0, z = 0;
            m.GetReportProperty("COG_X", ref x);
            m.GetReportProperty("COG_Y", ref y);
            m.GetReportProperty("COG_Z", ref z);
            return new Point(x, y, z);
        }

        public static string Prompt(string t, string msg, string def)
        {
            Form f = new Form() { Width = 300, Height = 150, Text = t, StartPosition = FormStartPosition.CenterScreen, TopMost = true };
            Label lbl = new Label() { Left = 10, Top = 15, Width = 275, Text = msg };
            TextBox txt = new TextBox() { Left = 50, Top = 40, Width = 200, Text = def };
            Button b = new Button() { Text = "OK", Left = 110, Top = 80, DialogResult = DialogResult.OK };
            f.Controls.Add(lbl); f.Controls.Add(txt); f.Controls.Add(b);
            return f.ShowDialog() == DialogResult.OK ? txt.Text : "";
        }
        #endregion
    }

    public class GroupSignature
    {
        public string GroupName { get; set; }
        public string PostProfile { get; set; }
        public string CapPlateProfile { get; set; }
        public string PostName { get; set; }
        public string CapPlateName { get; set; }
        public List<double> CornerDistances { get; set; }
        public double CapPlateVolume { get; set; }
        public int PartCount { get; set; }
        public double DistanceTolerance { get; set; }
    }
}
