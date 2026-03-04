using System;
using System.Collections;
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

            try
            {
                Model MyModel = new Model();
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
            catch (Exception ex)
            {
                connectionStatusPanel.BackColor = System.Drawing.Color.Red;
                statusLabel.Text = "Connection failed.";
                MessageBox.Show("Error connecting to Tekla: " + ex.Message, "Connection Error");
            }
        }

        private void createSignatureButton_Click(object sender, EventArgs e)
        {
            if (!MyModel.GetConnectionStatus())
            {
                MessageBox.Show("Not connected to a Tekla Structures model.", "Connection Error");
                return;
            }
            if (string.IsNullOrWhiteSpace(projectNumberTextBox.Text))
            {
                MessageBox.Show("Please enter a Project Number first.", "Input Required");
                return;
            }

            statusLabel.Text = "Select parts in the model...";
            this.Refresh();
            var picker = new Picker();
            ModelObjectEnumerator selectedObjects;
            try
            {
                selectedObjects = picker.PickObjects(Picker.PickObjectsEnum.PICK_N_PARTS,
                                                     "Select the sample group of parts");
            }
            catch (Exception)
            {
                statusLabel.Text = "Connected to: " + MyModel.GetInfo().ModelName;
                return;
            }
            if (selectedObjects.GetSize() < 2)
            {
                MessageBox.Show("Please select at least 2 parts.", "Selection Error");
                statusLabel.Text = "Connected to: " + MyModel.GetInfo().ModelName;
                return;
            }

            var partsData = new List<Tuple<Part, Tekla.Structures.Geometry3d.Point>>();
            while (selectedObjects.MoveNext())
            {
                if (selectedObjects.Current is Part part)
                {
                    partsData.Add(new Tuple<Part, Tekla.Structures.Geometry3d.Point>(part, GetPartCog(part)));
                }
            }

            var groupName = Prompt("Enter Signature Name",
                                   "Please enter a name for this group signature:",
                                   "Group_Signature");
            if (string.IsNullOrEmpty(groupName)) return;

            var signature = new GroupSignature
            {
                GroupName = groupName,
                PartCount = partsData.Count,
                Tolerance = Convert.ToDouble(toleranceTextBox.Text),
                Parts = partsData.Select(p => GetPartProperties(p.Item1)).ToList(),
                Distances = CalculateSortedDistances(partsData.Select(p => p.Item2).ToList())
            };

            var projectNumber = projectNumberTextBox.Text;
            var pluginDataBasePath = @"C:\TeklaGroupFinderData";
            var projectFolderPath = Path.Combine(pluginDataBasePath, projectNumber);
            Directory.CreateDirectory(projectFolderPath);
            var filePath = Path.Combine(projectFolderPath, $"{groupName}.json");
            File.WriteAllText(filePath, JsonConvert.SerializeObject(signature, Formatting.Indented));

            statusLabel.Text = "Connected to: " + MyModel.GetInfo().ModelName;
            MessageBox.Show($"Successfully created signature file:\n{filePath}", "Success");
        }

        private void findMatchesButton_Click(object sender, EventArgs e)
        {
            if (MyModel == null || !MyModel.GetConnectionStatus())
            {
                MessageBox.Show("Not connected to a Tekla Structures model.", "Connection Error");
                return;
            }

            MessageBox.Show("For this demonstration, please select a group of parts to form a sub-assembly.",
                            "Demonstration Step");

            var picker = new Picker();
            ModelObjectEnumerator aGroupToProcess;
            try
            {
                aGroupToProcess = picker.PickObjects(Picker.PickObjectsEnum.PICK_N_PARTS,
                                                     "Select parts to form a sub-assembly");
            }
            catch (Exception)
            {
                statusLabel.Text = "Connected to: " + MyModel.GetInfo().ModelName;
                return;
            }

            if (aGroupToProcess.GetSize() > 1 && createSubAssemblyCheckBox.Checked)
            {
                var partsToGroup = new ArrayList();
                Part mainPart = null;
                while (aGroupToProcess.MoveNext())
                {
                    if (aGroupToProcess.Current is Part part)
                    {
                        partsToGroup.Add(part);
                        if (mainPart == null) mainPart = part;
                    }
                }

                var parentAssembly = mainPart?.GetAssembly();
                if (parentAssembly == null)
                {
                    MessageBox.Show("Could not find the parent assembly.", "Error");
                    return;
                }

                var newSubAssembly = new Assembly { Name = "GROUP_ASSEMBLY" };
                newSubAssembly.SetMainPart(mainPart);
                for (int i = 1; i < partsToGroup.Count; i++)
                {
                    newSubAssembly.Add(partsToGroup[i] as Part);
                }

                if (!newSubAssembly.Insert())
                {
                    MessageBox.Show("Failed to create sub-assembly.", "Error");
                }
                else
                {
                    parentAssembly.Modify();
                }
            }

            MyModel.CommitChanges("Sub-assemblies created by Group Finder application.");
            statusLabel.Text = "Connected to: " + MyModel.GetInfo().ModelName;
            MessageBox.Show("Demonstration finished. Parts have been grouped.", "Process Complete");
        }

        private void openCatalogueButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The Catalogue will be built in a later step.", "Not Implemented");
        }

        #region Helper Methods
        private Tekla.Structures.Geometry3d.Point GetPartCog(Part part)
        {
            double cogX = 0, cogY = 0, cogZ = 0;
            part.GetReportProperty("COG_X", ref cogX);
            part.GetReportProperty("COG_Y", ref cogY);
            part.GetReportProperty("COG_Z", ref cogZ);
            return new Tekla.Structures.Geometry3d.Point(cogX, cogY, cogZ);
        }

        private Dictionary<string, string> GetPartProperties(Part part)
        {
            var props = new Dictionary<string, string>();
            string profile = "", grade = "";
            part.GetReportProperty("PROFILE", ref profile);
            part.GetReportProperty("MATERIAL_GRADE", ref grade);
            props.Add("Profile", profile);
            props.Add("Grade", grade);
            return props;
        }

        private List<double> CalculateSortedDistances(List<Tekla.Structures.Geometry3d.Point> cogs)
        {
            var distances = new List<double>();
            for (int i = 0; i < cogs.Count; i++)
            {
                for (int j = i + 1; j < cogs.Count; j++)
                {
                    distances.Add(Math.Round(
                        Tekla.Structures.Geometry3d.Distance.PointToPoint(cogs[i], cogs[j]), 2));
                }
            }
            distances.Sort();
            return distances;
        }

        public static string Prompt(string title, string promptText, string defaultValue)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = title,
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 50, Top = 20, Text = promptText, Width = 400 };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400, Text = defaultValue };
            Button confirmation = new Button()
            {
                Text = "Ok",
                Left = 350,
                Width = 100,
                Top = 80,
                DialogResult = DialogResult.OK
            };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;
            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }
        #endregion
    }

    public class GroupSignature
    {
        public string GroupName { get; set; }
        public int PartCount { get; set; }
        public List<Dictionary<string, string>> Parts { get; set; }
        public List<double> Distances { get; set; }
        public double Tolerance { get; set; }
    }
}