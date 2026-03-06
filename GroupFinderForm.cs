using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
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

            AddFabricationReportButton();
        }

        private void AddFabricationReportButton()
        {
            var fabBtn = new Button();
            fabBtn.Text = "Open Fabrication Report";
            fabBtn.Size = openCatalogueButton.Size;
            fabBtn.Location = new System.Drawing.Point(openCatalogueButton.Left, openCatalogueButton.Bottom + 10);
            fabBtn.Click += (s, e) =>
            {
                var path = GetDataPath();
                using (var fabForm = new FabricationReportForm(path, udaNameTextBox.Text))
                {
                    fabForm.ShowDialog(this);
                }
            };
            this.Controls.Add(fabBtn);

            foreach (Control c in this.Controls)
            {
                if (c != fabBtn && c.Top > openCatalogueButton.Top)
                {
                    c.Top += (fabBtn.Height + 10);
                }
            }
            this.Height += (fabBtn.Height + 10);
        }

        private void createSignatureButton_Click(object sender, EventArgs e)
        {
            if (MyModel == null || !MyModel.GetConnectionStatus()) return;

            var picker = new Picker();
            var selected = picker.PickObjects(Picker.PickObjectsEnum.PICK_N_PARTS, "Select Post and Cap Plate");
            var parts = new List<Part>();
            while (selected.MoveNext()) if (selected.Current is Part p) parts.Add(p);

            if (parts.Count != 2)
            {
                MessageBox.Show("Selection must include exactly two parts: one post and one cap plate.", "Selection Error");
                return;
            }

            double vol0 = 0, vol1 = 0;
            parts[0].GetReportProperty("VOLUME", ref vol0);
            parts[1].GetReportProperty("VOLUME", ref vol1);

            Part mainPost = vol0 >= vol1 ? parts[0] : parts[1];
            Part capPlate = vol0 >= vol1 ? parts[1] : parts[0];

            var groupName = Prompt("New Signature", "Enter Signature Name:", "Type_1");
            if (string.IsNullOrEmpty(groupName)) return;

            double.TryParse(toleranceTextBox.Text, out double distTolerance);

            GetPostEndpoints(mainPost, out Point startPt, out Point endPt);
            Point capCog = GetPartCog(capPlate);
            Point refNode = Distance.PointToPoint(startPt, capCog) <= Distance.PointToPoint(endPt, capCog) ? startPt : endPt;

            double volume = Math.Min(vol0, vol1);

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
            public Part Plate { get; set; }
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
            var rawPlates = new List<Part>();
            var iter = MyModel.GetModelObjectSelector().GetAllObjectsWithType(new Type[] { typeof(Part) });

            while (iter.MoveNext())
            {
                if (iter.Current is Part part)
                {
                    string profileKey = part.Profile.ProfileString.Trim().ToUpper();
                    string nameKey = part.Name?.Trim().ToUpper() ?? "";

                    bool isPlate = validPlateProfiles.Contains(profileKey) && (validPlateNames.Count == 0 || validPlateNames.Contains(nameKey));
                    bool isPost = validPostProfiles.Contains(profileKey) && (validPostNames.Count == 0 || validPostNames.Contains(nameKey));

                    if (isPlate) rawPlates.Add(part);
                    if (isPost)
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
                }
            });

            MyModel.CommitChanges("Posts Grouped.");
            processingProgressBar.Visible = false;
            statusLabel.Text = $"Tagged {matchCount} posts.";
            MessageBox.Show($"Search complete. {matchCount} matches found.", "Result");
        }

        private void openCatalogueButton_Click(object sender, EventArgs e)
        {
            var path = GetDataPath();
            if (!Directory.Exists(path) || Directory.GetFiles(path, "*.json").Length == 0)
            {
                MessageBox.Show("No signatures have been created yet.", "Catalogue Empty", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var catalogueForm = new SignatureCatalogueForm(path, udaNameTextBox.Text))
            {
                catalogueForm.ShowDialog(this);
            }
        }

        #region Helper Methods
        public static bool DoSignaturesMatch(GroupSignature s, string postN, string postP, string plateN, string plateP, double volume, List<double> c, double t)
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

        public static Point GetReportPoint(Part p, string prefix)
        {
            double x = 0, y = 0, z = 0;
            p.GetReportProperty(prefix + "_X", ref x);
            p.GetReportProperty(prefix + "_Y", ref y);
            p.GetReportProperty(prefix + "_Z", ref z);
            return new Point(x, y, z);
        }

        public static void GetPostEndpoints(Part post, out Point start, out Point end)
        {
            if (post is Beam beam) { start = beam.StartPoint; end = beam.EndPoint; }
            else { start = GetReportPoint(post, "START"); end = GetReportPoint(post, "END"); }
        }

        public static List<double> GetCornerDistances(Part p, Point r)
        {
            var d = new List<double>();
            if (p is ContourPlate cp)
            {
                foreach (ContourPoint pt in cp.Contour.ContourPoints)
                    d.Add(Math.Round(Distance.PointToPoint(r, new Point(pt.X, pt.Y, pt.Z)), 1));
            }
            else if (p is Beam b)
            {
                d.Add(Math.Round(Distance.PointToPoint(r, b.StartPoint), 1));
                d.Add(Math.Round(Distance.PointToPoint(r, b.EndPoint), 1));
            }
            d.Sort();
            return d;
        }

        public static double DistanceSquared(Point p1, Point p2)
        {
            double dx = p1.X - p2.X;
            double dy = p1.Y - p2.Y;
            double dz = p1.Z - p2.Z;
            return (dx * dx) + (dy * dy) + (dz * dz);
        }

        public static Point GetPartCog(ModelObject m)
        {
            double x = 0, y = 0, z = 0;
            if (m.GetReportProperty("COG_X", ref x) &&
                m.GetReportProperty("COG_Y", ref y) &&
                m.GetReportProperty("COG_Z", ref z))
            {
                return new Point(x, y, z);
            }

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
            return new Point();
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

    public class SignatureCatalogueForm : Form
    {
        private string DataPath;
        private string UdaName;
        private DataGridView Grid;
        private Model MyModel;

        public SignatureCatalogueForm(string path, string udaName)
        {
            DataPath = path;
            UdaName = udaName;
            MyModel = new Model();
            InitializeComponent();

            this.Shown += SignatureCatalogueForm_Shown;
        }

        private void InitializeComponent()
        {
            this.Text = "Signature Catalogue";
            this.Width = 900;
            this.Height = 450;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowIcon = false;

            Grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = System.Drawing.Color.White,
                RowHeadersVisible = false
            };

            var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
            var deleteBtn = new Button { Text = "Delete Selected", Width = 120, Height = 30, Left = 10, Top = 10 };
            deleteBtn.Click += DeleteBtn_Click;
            var selectBtn = new Button { Text = "Select in Model", Width = 120, Height = 30, Left = 140, Top = 10 };
            selectBtn.Click += SelectBtn_Click;
            var closeBtn = new Button { Text = "Close", Width = 100, Height = 30, Left = 270, Top = 10 };
            closeBtn.Click += (s, e) => this.Close();

            bottomPanel.Controls.Add(deleteBtn);
            bottomPanel.Controls.Add(selectBtn);
            bottomPanel.Controls.Add(closeBtn);
            this.Controls.Add(Grid);
            this.Controls.Add(bottomPanel);
        }

        private async void SignatureCatalogueForm_Shown(object sender, EventArgs e)
        {
            this.Text = "Signature Catalogue - LOADING DATA...";

            object gridData = null;
            await System.Threading.Tasks.Task.Run(() =>
            {
                gridData = LoadSignaturesData();
            });

            if (gridData != null)
            {
                Grid.DataSource = gridData;
                if (Grid.Columns["FilePath"] != null) Grid.Columns["FilePath"].Visible = false;
            }

            this.Text = "Signature Catalogue";
        }

        private Dictionary<string, int> GetModelGroupCounts()
        {
            var counts = new Dictionary<string, int>();
            if (MyModel == null || !MyModel.GetConnectionStatus()) return counts;

            var iter = MyModel.GetModelObjectSelector().GetAllObjectsWithType(new Type[] { typeof(Part) });
            while (iter.MoveNext())
            {
                if (iter.Current is Part p)
                {
                    string udaVal = "";
                    if (p.GetUserProperty(UdaName, ref udaVal) && !string.IsNullOrEmpty(udaVal))
                    {
                        if (!counts.ContainsKey(udaVal)) counts[udaVal] = 0;
                        counts[udaVal]++;
                    }
                }
            }

            var groupCounts = new Dictionary<string, int>();
            foreach (var kvp in counts) groupCounts[kvp.Key] = kvp.Value / 2;
            return groupCounts;
        }

        private object LoadSignaturesData()
        {
            if (!Directory.Exists(DataPath)) return null;
            var counts = GetModelGroupCounts();
            var files = Directory.GetFiles(DataPath, "*.json");

            return files.Select(f =>
            {
                var sig = JsonConvert.DeserializeObject<GroupSignature>(File.ReadAllText(f));
                return new
                {
                    FilePath = f,
                    Group = sig.GroupName,
                    QtyFound = counts.ContainsKey(sig.GroupName) ? counts[sig.GroupName] : 0,
                    Post = sig.PostProfile,
                    Plate = sig.CapPlateProfile,
                    Volume = Math.Round(sig.CapPlateVolume, 1)
                };
            }).ToList();
        }

        private void SelectBtn_Click(object sender, EventArgs e)
        {
            if (Grid.SelectedRows.Count == 0) return;
            var selectedGroups = new HashSet<string>();
            foreach (DataGridViewRow row in Grid.SelectedRows) selectedGroups.Add(row.Cells["Group"].Value?.ToString());

            var objectsToSelect = new ArrayList();
            var iter = MyModel.GetModelObjectSelector().GetAllObjectsWithType(new Type[] { typeof(Part) });
            while (iter.MoveNext())
            {
                if (iter.Current is Part p)
                {
                    string udaVal = "";
                    if (p.GetUserProperty(UdaName, ref udaVal) && selectedGroups.Contains(udaVal))
                        objectsToSelect.Add(p);
                }
            }

            if (objectsToSelect.Count > 0) new Tekla.Structures.Model.UI.ModelObjectSelector().Select(objectsToSelect);
            else MessageBox.Show("No parts found.", "Select", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private async void DeleteBtn_Click(object sender, EventArgs e)
        {
            if (Grid.SelectedRows.Count == 0) return;
            if (MessageBox.Show("Delete selected signature(s)?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in Grid.SelectedRows)
                {
                    var filePath = row.Cells["FilePath"].Value?.ToString();
                    if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath)) File.Delete(filePath);
                }

                this.Text = "Signature Catalogue - REFRESHING...";
                object gridData = null;
                await System.Threading.Tasks.Task.Run(() =>
                {
                    gridData = LoadSignaturesData();
                });

                if (gridData != null) Grid.DataSource = gridData;
                this.Text = "Signature Catalogue";
            }
        }
    }

    public class FabricationReportForm : Form
    {
        private string DataPath;
        private string UdaName;
        private DataGridView Grid;
        private TextBox SearchBox;
        private Button ExportBtn, SelectBtn, ApplyBtn, CloseBtn;

        private CheckBox chk1, chk2, chk3, chk4, chk5;

        public FabricationReportForm(string path, string udaName)
        {
            DataPath = path;
            UdaName = udaName;
            InitializeComponent();

            this.Shown += FabricationReportForm_Shown;
        }

        private void InitializeComponent()
        {
            this.Text = "Fabrication Report";
            this.Width = 1200;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowIcon = false;

            var topPanel = new Panel { Dock = DockStyle.Top, Height = 50 };

            var searchLabel = new Label { Text = "Search/Filter:", Left = 10, Top = 15, Width = 80 };
            SearchBox = new TextBox { Left = 90, Top = 12, Width = 150 };
            SearchBox.TextChanged += (s, e) => ApplyFilters();

            var pennantLabel = new Label { Text = "Pennants:", Left = 260, Top = 15, Width = 60 };
            chk1 = new CheckBox { Text = "1", Checked = true, Left = 320, Top = 14, Width = 35 };
            chk2 = new CheckBox { Text = "2", Checked = true, Left = 355, Top = 14, Width = 35 };
            chk3 = new CheckBox { Text = "3", Checked = true, Left = 390, Top = 14, Width = 35 };
            chk4 = new CheckBox { Text = "4", Checked = true, Left = 425, Top = 14, Width = 35 };
            chk5 = new CheckBox { Text = "5", Checked = true, Left = 460, Top = 14, Width = 35 };

            chk1.CheckedChanged += (s, e) => ApplyFilters();
            chk2.CheckedChanged += (s, e) => ApplyFilters();
            chk3.CheckedChanged += (s, e) => ApplyFilters();
            chk4.CheckedChanged += (s, e) => ApplyFilters();
            chk5.CheckedChanged += (s, e) => ApplyFilters();

            topPanel.Controls.Add(searchLabel);
            topPanel.Controls.Add(SearchBox);
            topPanel.Controls.Add(pennantLabel);
            topPanel.Controls.Add(chk1);
            topPanel.Controls.Add(chk2);
            topPanel.Controls.Add(chk3);
            topPanel.Controls.Add(chk4);
            topPanel.Controls.Add(chk5);

            Grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = System.Drawing.Color.White,
                RowHeadersVisible = false
            };

            var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };

            ExportBtn = new Button { Text = "Export to CSV (Excel)", Width = 150, Height = 30, Left = 10, Top = 10 };
            ExportBtn.Click += ExportBtn_Click;

            SelectBtn = new Button { Text = "Select in Model", Width = 120, Height = 30, Left = 170, Top = 10 };
            SelectBtn.Click += SelectBtn_Click;

            ApplyBtn = new Button { Text = "Write X Codes to Model", Width = 160, Height = 30, Left = 300, Top = 10 };
            ApplyBtn.Click += ApplyXCodesBtn_Click;

            CloseBtn = new Button { Text = "Close", Width = 100, Height = 30, Left = 470, Top = 10 };
            CloseBtn.Click += (s, e) => this.Close();

            bottomPanel.Controls.Add(ExportBtn);
            bottomPanel.Controls.Add(SelectBtn);
            bottomPanel.Controls.Add(ApplyBtn);
            bottomPanel.Controls.Add(CloseBtn);

            this.Controls.Add(Grid);
            this.Controls.Add(topPanel);
            this.Controls.Add(bottomPanel);
        }

        private async void FabricationReportForm_Shown(object sender, EventArgs e)
        {
            this.Text = "Fabrication Report - LOADING DATA, PLEASE WAIT...";
            SearchBox.Enabled = false;
            ExportBtn.Enabled = false;
            SelectBtn.Enabled = false;
            ApplyBtn.Enabled = false;

            DataTable dt = null;
            await System.Threading.Tasks.Task.Run(() =>
            {
                dt = GenerateReportData();
            });

            if (dt != null)
            {
                Grid.DataSource = dt.DefaultView;
                if (Grid.Columns["POST_PARTS"] != null)
                    Grid.Columns["POST_PARTS"].Visible = false;
                ApplyFilters();
            }

            SearchBox.Enabled = true;
            ExportBtn.Enabled = true;
            SelectBtn.Enabled = true;
            ApplyBtn.Enabled = true;
            this.Text = "Fabrication Report";
        }

        private string FormatToArchitectural(double roundedMmForAggregation)
        {
            double totalInchesDecimal = roundedMmForAggregation / 25.4;
            int feet = (int)(totalInchesDecimal / 12.0);
            double remainingInches = totalInchesDecimal - (feet * 12.0);

            int wholeInches = (int)Math.Floor(remainingInches);
            int numerator = (int)Math.Round((remainingInches - wholeInches) * 16.0);

            if (numerator == 16)
            {
                wholeInches++;
                numerator = 0;
            }
            if (wholeInches == 12)
            {
                feet++;
                wholeInches = 0;
            }

            string fractionStr = "";
            if (numerator > 0)
            {
                int denom = 16;
                while (numerator % 2 == 0 && denom % 2 == 0)
                {
                    numerator /= 2;
                    denom /= 2;
                }
                fractionStr = $"{numerator}/{denom}";
            }

            if (feet > 0)
            {
                if (wholeInches == 0 && fractionStr == "") return $"{feet}'-0\"";
                if (wholeInches == 0) return $"{feet}'-0 {fractionStr}\"";
                if (fractionStr == "") return $"{feet}'-{wholeInches}\"";
                return $"{feet}'-{wholeInches} {fractionStr}\"";
            }
            else
            {
                if (wholeInches == 0 && fractionStr == "") return "0\"";
                if (wholeInches == 0) return $"{fractionStr}\"";
                if (fractionStr == "") return $"{wholeInches}\"";
                return $"{wholeInches} {fractionStr}\"";
            }
        }

        private string FormatMetricPlateProfileToImperial(string profileString)
        {
            if (string.IsNullOrEmpty(profileString)) return "";

            if (profileString.Contains("\"") || profileString.Contains("/")) return profileString.ToUpper();

            string upper = profileString.ToUpper().Trim();

            if (!upper.StartsWith("PL") && !upper.StartsWith("FL") && !upper.StartsWith("PLT")) return upper;

            int firstDigitIdx = -1;
            for (int i = 0; i < upper.Length; i++)
            {
                if (char.IsDigit(upper[i]) || upper[i] == '.')
                {
                    firstDigitIdx = i;
                    break;
                }
            }

            if (firstDigitIdx == -1) return upper;

            string prefix = upper.Substring(0, firstDigitIdx).Trim();
            string rest = upper.Substring(firstDigitIdx);

            char[] separators = new char[] { 'X', '*' };
            string[] parts = rest.Split(separators);

            List<string> formattedParts = new List<string>();
            foreach (var p in parts)
            {
                if (double.TryParse(p.Trim(), out double mmVal))
                {
                    formattedParts.Add(FormatToArchitectural(mmVal));
                }
                else
                {
                    formattedParts.Add(p.Trim());
                }
            }

            return prefix + string.Join("X", formattedParts);
        }

        private DataTable GenerateReportData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("POST PENNANT", typeof(string));
            dt.Columns.Add("POST GROUP TYPE", typeof(string));
            dt.Columns.Add("POST OVERALL LENGTH", typeof(string));
            dt.Columns.Add("POST PROFILE", typeof(string));
            dt.Columns.Add("POST LENGTH", typeof(string));
            dt.Columns.Add("PLATE PROFILE", typeof(string));
            dt.Columns.Add("PLATE THICKNESS", typeof(string));
            dt.Columns.Add("QUANTITY", typeof(int));
            dt.Columns.Add("X CODE", typeof(string));
            dt.Columns.Add("POST_PARTS", typeof(object));

            var myModel = new Model();
            if (!myModel.GetConnectionStatus() || !Directory.Exists(DataPath)) return dt;

            var signatures = Directory.GetFiles(DataPath, "*.json")
                .Select(f => JsonConvert.DeserializeObject<GroupSignature>(File.ReadAllText(f)))
                .ToDictionary(s => s.GroupName, s => s);

            var partsByUda = new Dictionary<string, List<Part>>();
            var iter = myModel.GetModelObjectSelector().GetAllObjectsWithType(new Type[] { typeof(Part) });

            while (iter.MoveNext())
            {
                if (iter.Current is Part p)
                {
                    string udaVal = "";
                    if (p.GetUserProperty(UdaName, ref udaVal) && !string.IsNullOrEmpty(udaVal))
                    {
                        if (!partsByUda.ContainsKey(udaVal)) partsByUda[udaVal] = new List<Part>();
                        partsByUda[udaVal].Add(p);
                    }
                }
            }

            var reportRows = new List<ReportItem>();

            foreach (var kvp in partsByUda)
            {
                string groupName = kvp.Key;
                if (!signatures.ContainsKey(groupName)) continue;

                var sig = signatures[groupName];
                var partsInGroup = kvp.Value;

                double plateThickness = 0;
                var plates = partsInGroup.Where(p => p.Profile.ProfileString.Trim().ToUpper() == sig.CapPlateProfile.Trim().ToUpper()).ToList();
                var posts = partsInGroup.Where(p => p.Profile.ProfileString.Trim().ToUpper() == sig.PostProfile.Trim().ToUpper()).ToList();

                if (plates.Count > 0)
                {
                    var samplePlate = plates.First();
                    if (!samplePlate.GetReportProperty("PROFILE.PLATE_THICKNESS", ref plateThickness) || plateThickness == 0)
                    {
                        double w = 0, h = 0;
                        samplePlate.GetReportProperty("PROFILE.WIDTH", ref w);
                        samplePlate.GetReportProperty("PROFILE.HEIGHT", ref h);
                        plateThickness = Math.Min(w > 0 ? w : double.MaxValue, h > 0 ? h : double.MaxValue);
                        if (plateThickness == double.MaxValue) plateThickness = 0;
                    }
                }

                var plateCogs = plates.Select(pl => new { Plate = pl, Cog = GroupFinderForm.GetPartCog(pl) }).ToList();

                foreach (var post in posts)
                {
                    Phase phase;
                    int phaseNum = 0;
                    if (post.GetPhase(out phase))
                    {
                        phaseNum = phase.PhaseNumber;
                    }

                    string pennant = Math.Floor(phaseNum / 100.0).ToString();

                    double length = 0;
                    post.GetReportProperty("LENGTH", ref length);

                    double overallLengthMm = length + plateThickness;

                    double decimalInches = Math.Round((overallLengthMm / 25.4) * 16.0) / 16.0;
                    double roundedMmForAggregation = decimalInches * 25.4;

                    string formattedOverallLength = FormatToArchitectural(roundedMmForAggregation);
                    string formattedPostLength = FormatToArchitectural(length);
                    string formattedPlateThickness = FormatToArchitectural(plateThickness);

                    string postProfile = post.Profile.ProfileString;
                    string plateProfile = FormatMetricPlateProfileToImperial(sig.CapPlateProfile);

                    string typeLetter = groupName.Contains("_") ? groupName.Substring(groupName.LastIndexOf('_') + 1) : groupName;
                    string xCode = $"X{typeLetter}{decimalInches.ToString("0.000")}";

                    GroupFinderForm.GetPostEndpoints(post, out Point startPt, out Point endPt);
                    Part matchedPlate = null;
                    double minSq = double.MaxValue;

                    foreach (var pc in plateCogs)
                    {
                        double d1 = GroupFinderForm.DistanceSquared(pc.Cog, startPt);
                        double d2 = GroupFinderForm.DistanceSquared(pc.Cog, endPt);
                        double d = Math.Min(d1, d2);
                        if (d < minSq)
                        {
                            minSq = d;
                            matchedPlate = pc.Plate;
                        }
                    }

                    var linkedParts = new List<Part> { post };
                    if (matchedPlate != null && minSq < (300.0 * 300.0))
                    {
                        linkedParts.Add(matchedPlate);
                    }

                    reportRows.Add(new ReportItem
                    {
                        Pennant = pennant,
                        GroupType = groupName,
                        OverallLengthSortKey = roundedMmForAggregation,
                        FormattedOverallLength = formattedOverallLength,
                        PostProfile = postProfile,
                        PostLength = formattedPostLength,
                        PlateProfile = plateProfile,
                        PlateThickness = formattedPlateThickness,
                        XCode = xCode,
                        LinkedParts = linkedParts
                    });
                }
            }

            var aggregatedData = reportRows
                .GroupBy(r => new {
                    r.Pennant,
                    r.GroupType,
                    r.FormattedOverallLength,
                    r.PostProfile,
                    r.PostLength,
                    r.PlateProfile,
                    r.PlateThickness,
                    r.OverallLengthSortKey,
                    r.XCode
                })
                .Select(g => new {
                    g.Key.Pennant,
                    g.Key.GroupType,
                    g.Key.FormattedOverallLength,
                    g.Key.PostProfile,
                    g.Key.PostLength,
                    g.Key.PlateProfile,
                    g.Key.PlateThickness,
                    g.Key.OverallLengthSortKey,
                    Quantity = g.Count(),
                    g.Key.XCode,
                    PartsToSelect = g.SelectMany(x => x.LinkedParts).ToList()
                })
                .OrderBy(r => r.Pennant).ThenBy(r => r.GroupType).ThenBy(r => r.OverallLengthSortKey);

            foreach (var item in aggregatedData)
            {
                dt.Rows.Add(
                    item.Pennant,
                    item.GroupType,
                    item.FormattedOverallLength,
                    item.PostProfile,
                    item.PostLength,
                    item.PlateProfile,
                    item.PlateThickness,
                    item.Quantity,
                    item.XCode,
                    item.PartsToSelect
                );
            }

            return dt;
        }

        private void ApplyFilters()
        {
            if (Grid.DataSource is DataView dv)
            {
                var allowedPennants = new List<string>();
                if (chk1.Checked) allowedPennants.Add("'1'");
                if (chk2.Checked) allowedPennants.Add("'2'");
                if (chk3.Checked) allowedPennants.Add("'3'");
                if (chk4.Checked) allowedPennants.Add("'4'");
                if (chk5.Checked) allowedPennants.Add("'5'");

                string pennantFilter = allowedPennants.Count > 0
                    ? $"[POST PENNANT] IN ({string.Join(",", allowedPennants)})"
                    : "[POST PENNANT] IN ('-1')";

                string q = SearchBox.Text.Replace("'", "''");
                string searchFilter = string.IsNullOrWhiteSpace(q)
                    ? ""
                    : $"([POST PENNANT] LIKE '%{q}%' OR [POST GROUP TYPE] LIKE '%{q}%' OR [POST PROFILE] LIKE '%{q}%' OR [PLATE PROFILE] LIKE '%{q}%' OR [X CODE] LIKE '%{q}%')";

                if (string.IsNullOrEmpty(searchFilter))
                {
                    dv.RowFilter = pennantFilter;
                }
                else
                {
                    dv.RowFilter = $"({pennantFilter}) AND ({searchFilter})";
                }
            }
        }

        private void SelectBtn_Click(object sender, EventArgs e)
        {
            if (Grid.SelectedRows.Count == 0) return;

            var objectsToSelect = new ArrayList();
            foreach (DataGridViewRow row in Grid.SelectedRows)
            {
                if (row.DataBoundItem is DataRowView drv && drv["POST_PARTS"] is List<Part> linkedParts)
                {
                    foreach (var p in linkedParts)
                    {
                        objectsToSelect.Add(p);
                    }
                }
            }

            if (objectsToSelect.Count > 0)
            {
                new Tekla.Structures.Model.UI.ModelObjectSelector().Select(objectsToSelect);
            }
            else
            {
                MessageBox.Show("No parts found to select.", "Select", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void ApplyXCodesBtn_Click(object sender, EventArgs e)
        {
            var myModel = new Model();
            if (!myModel.GetConnectionStatus()) return;

            this.Text = "Fabrication Report - WRITING TO MODEL, PLEASE WAIT...";
            ApplyBtn.Enabled = false;

            int modifiedCount = 0;

            var updates = new List<Tuple<string, List<Part>>>();
            if (Grid.DataSource is DataView dv && dv.Table != null)
            {
                foreach (DataRow row in dv.Table.Rows)
                {
                    string xCode = row["X CODE"]?.ToString();
                    if (!string.IsNullOrEmpty(xCode) && row["POST_PARTS"] is List<Part> linkedParts)
                    {
                        updates.Add(new Tuple<string, List<Part>>(xCode, linkedParts));
                    }
                }
            }

            await System.Threading.Tasks.Task.Run(() =>
            {
                foreach (var update in updates)
                {
                    foreach (var part in update.Item2)
                    {
                        part.SetUserProperty("HERRICK_CODE", "X");
                        part.SetUserProperty("BOM_Remarks", "P_POST_" + update.Item1);
                        part.Modify();
                        modifiedCount++;
                    }
                }
            });

            if (modifiedCount > 0)
            {
                myModel.CommitChanges("Applied X Codes to HERRICK_CODE and BOM_Remarks");
                MessageBox.Show($"Successfully applied UDAs to {modifiedCount} parts in the model.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("No parts were found to modify.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            this.Text = "Fabrication Report";
            ApplyBtn.Enabled = true;
        }

        private void ExportBtn_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "CSV Excel File (*.csv)|*.csv", FileName = "FabricationReport.csv" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var sb = new System.Text.StringBuilder();

                    var visibleCols = Grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();

                    var headers = visibleCols.Select(c => "\"" + c.HeaderText.Replace("\"", "\"\"") + "\"");
                    sb.AppendLine(string.Join(",", headers));

                    foreach (DataGridViewRow row in Grid.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            var cells = visibleCols.Select(c =>
                            {
                                string val = row.Cells[c.Index].Value?.ToString() ?? "";
                                // Standard CSV escape: duplicate quotes inside the string, then wrap the whole string in quotes
                                return "\"" + val.Replace("\"", "\"\"") + "\"";
                            });
                            sb.AppendLine(string.Join(",", cells));
                        }
                    }

                    File.WriteAllText(sfd.FileName, sb.ToString());
                    MessageBox.Show("Exported Successfully! You can open this file directly in Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private class ReportItem
        {
            public string Pennant { get; set; }
            public string GroupType { get; set; }
            public double OverallLengthSortKey { get; set; }
            public string FormattedOverallLength { get; set; }
            public string PostProfile { get; set; }
            public string PostLength { get; set; }
            public string PlateProfile { get; set; }
            public string PlateThickness { get; set; }
            public string XCode { get; set; }
            public List<Part> LinkedParts { get; set; }
        }
    }
}
