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
            Point refNode = GetPostTopNode(mainPost);

            var signature = new GroupSignature
            {
                GroupName = groupName,
                PostProfile = GetProperty(mainPost, "PROFILE"),
                CapPlateProfile = GetProperty(capPlate, "PROFILE"),
                CornerDistances = GetCornerDistances(capPlate, refNode),
                VoidDistances = GetVoidDistances(capPlate, refNode),
                PartCount = parts.Count,
                DistanceTolerance = distTolerance > 0 ? distTolerance : 2.0
            };

            var path = GetDataPath();
            Directory.CreateDirectory(path);
            File.WriteAllText(Path.Combine(path, $"{groupName}.json"), JsonConvert.SerializeObject(signature, Formatting.Indented));

            statusLabel.Text = $"Signature '{groupName}' Created.";
        }

        private async void findMatchesButton_Click(object sender, EventArgs e)
        {
            if (MyModel == null || !MyModel.GetConnectionStatus()) return;
            var path = GetDataPath();

            if (!Directory.Exists(path)) return;

            var signatures = Directory.GetFiles(path, "*.json")
                .Select(f => JsonConvert.DeserializeObject<GroupSignature>(File.ReadAllText(f))).ToList();

            var targetAssemblies = new List<Assembly>();
            var iter = MyModel.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.ASSEMBLY);
            while (iter.MoveNext())
            {
                if (iter.Current is Assembly assy)
                {
                    bool hasCapPlate = false;
                    foreach (object sec in assy.GetSecondaries())
                    {
                        if (sec is ContourPlate) { hasCapPlate = true; break; }
                    }
                    if (hasCapPlate) targetAssemblies.Add(assy);
                }
            }

            processingProgressBar.Visible = true;
            processingProgressBar.Maximum = targetAssemblies.Count;
            int matchCount = 0;
            double.TryParse(toleranceTextBox.Text, out double globalTol);

            await System.Threading.Tasks.Task.Run(() =>
            {
                for (int i = 0; i < targetAssemblies.Count; i++)
                {
                    this.Invoke((MethodInvoker)delegate { processingProgressBar.Value = i + 1; });
                    Assembly assy = targetAssemblies[i];
                    Part main = assy.GetMainPart() as Part;
                    ContourPlate plate = null;
                    foreach (object sec in assy.GetSecondaries()) if (sec is ContourPlate cp) { plate = cp; break; }

                    if (main == null || plate == null) continue;

                    string postProf = GetProperty(main, "PROFILE");
                    string plateProf = GetProperty(plate, "PROFILE");
                    Point refNode = GetPostTopNode(main);
                    var corners = GetCornerDistances(plate, refNode);
                    var voids = GetVoidDistances(plate, refNode);

                    foreach (var sig in signatures)
                    {
                        double activeTol = sig.DistanceTolerance > 0 ? sig.DistanceTolerance : globalTol;

                        if (DoSignaturesMatch(sig, postProf, plateProf, corners, voids, activeTol))
                        {
                            main.SetUserProperty(udaNameTextBox.Text, sig.GroupName);
                            main.Modify();
                            matchCount++;
                            break;
                        }
                    }
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
        private bool DoSignaturesMatch(GroupSignature s, string postP, string plateP, List<double> c, List<double> v, double t)
        {
            if (s.PostProfile.Trim().ToUpper() != postP.Trim().ToUpper() ||
                s.CapPlateProfile.Trim().ToUpper() != plateP.Trim().ToUpper()) return false;

            if (s.CornerDistances.Count != c.Count || s.VoidDistances.Count != v.Count) return false;

            for (int i = 0; i < c.Count; i++)
                if (Math.Abs(s.CornerDistances[i] - c[i]) > t) return false;

            for (int i = 0; i < v.Count; i++)
                if (Math.Abs(s.VoidDistances[i] - v[i]) > t) return false;

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

        private Point GetPostTopNode(Part post)
        {
            var start = GetReportPoint(post, "START");
            var end = GetReportPoint(post, "END");
            return end.Z >= start.Z ? end : start;
        }

        private List<double> GetCornerDistances(ContourPlate p, Point r)
        {
            var d = new List<double>();
            foreach (ContourPoint cp in p.Contour.ContourPoints)
                d.Add(Math.Round(Distance.PointToPoint(r, new Point(cp.X, cp.Y, cp.Z)), 1));
            d.Sort(); return d;
        }

        private List<double> GetVoidDistances(Part p, Point r)
        {
            var d = new List<double>();
            var bools = p.GetBooleans();
            while (bools.MoveNext()) if (bools.Current is BooleanPart bp) d.Add(Math.Round(Distance.PointToPoint(r, GetPartCog(bp.OperativePart)), 1));
            var bolts = p.GetBolts();
            while (bolts.MoveNext()) if (bolts.Current is BoltGroup bg) foreach (Point pos in bg.BoltPositions) d.Add(Math.Round(Distance.PointToPoint(r, pos), 1));
            d.Sort(); return d;
        }

        private Point GetPartCog(ModelObject m)
        {
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
        public List<double> CornerDistances { get; set; }
        public List<double> VoidDistances { get; set; }
        public int PartCount { get; set; }
        public double DistanceTolerance { get; set; }
    }
}