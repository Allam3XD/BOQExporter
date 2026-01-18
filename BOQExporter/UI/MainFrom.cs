using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using BOQExporter.Revit.Extraction;
using BOQExporter.Revit.Modeles;
using BOQExporter.Revit.Utils;
using static BOQExporter.Revit.Utils.RvtUtils.CategoryUtils;

namespace BOQExporter.UI
{
    
    //System.Windows.Forms.Form
    public partial class MainFrom : MetroFramework.Forms.MetroForm
    {
        public MainFrom()
        {

            
            InitializeComponent();
            // populate the category checklist box
            PopulateCategoryChecklist();


            // load previous selections if any
            LoadSelections();
            // load default project info from Revit document
            LoadProjectInfoDefaults();
            this.FormClosing += MainFrom_FormClosing;

        }


        /// <summary>
        /// fills the category checklist box with supported categories
        /// </summary>
        private void PopulateCategoryChecklist()
        {
            clbCategories.Items.Clear();

            foreach (var bic in MasterFormatMapper.SupportedCategories())
            {
                var item = new CategoryItem(bic);

                // by default , select all categories
                bool defaultChecked = (bic == null);

                
                clbCategories.Items.Add(item, defaultChecked);
            }
        }

        /// <summary>
        /// retrieves the selected categories from the checklist box
        /// </summary>
        /// <returns></returns>
        private HashSet<BuiltInCategory> GetSelectedCategories()
        {
            var selected = new HashSet<BuiltInCategory>();
            foreach (CategoryItem item in clbCategories.CheckedItems)
            {
                selected.Add(item.Bic);
            }
            return selected;
        }

        /// <summary>
        /// updates the label that shows the count of selected categories
        /// </summary>
        private void UpdateSelectedItemsLabel()
        {
            SelectedItems.Text = $"Selected Categories : {clbCategories.CheckedItems.Count}";
        }

        /// <summary>
        /// loads previously saved selections from SelectionMemory
        /// </summary>
        private void LoadSelections()
        {
            // Restore categories first
            foreach (var bic in SelectionMemory.SavedCategories)
            {
                for (int i = 0; i < clbCategories.Items.Count; i++)
                {
                    var item = (CategoryItem)clbCategories.Items[i];
                    if (item.Bic == bic)
                        clbCategories.SetItemChecked(i, true);
                }
            }

            // Restore group checkboxes visually, without firing ToggleGroup
            ArchCB.CheckedChanged -= ArchCB_CheckedChanged;
            StruCB.CheckedChanged -= StruCB_CheckedChanged_1;
            MechCB.CheckedChanged -= MechCB_CheckedChanged;
            ElecCB.CheckedChanged -= ElecCB_CheckedChanged_1;

            ArchCB.Checked = SelectionMemory.ArchChecked;
            StruCB.Checked = SelectionMemory.StruChecked;
            MechCB.Checked = SelectionMemory.MechChecked;
            ElecCB.Checked = SelectionMemory.ElecChecked;

            ArchCB.CheckedChanged += ArchCB_CheckedChanged;
            StruCB.CheckedChanged += StruCB_CheckedChanged_1;
            MechCB.CheckedChanged += MechCB_CheckedChanged;
            ElecCB.CheckedChanged += ElecCB_CheckedChanged_1;

            UpdateSelectedItemsLabel();

            // restore meta data fields
            txtProjectName.Text = SelectionMemory.ProjectName;
            txtClientName.Text = SelectionMemory.ClientName;
            txtLocation.Text = SelectionMemory.Location;
            dtpIssueDate.Text = SelectionMemory.IssueDate;

            if (DateTime.TryParse(SelectionMemory.IssueDate, out DateTime parsedDate))
            {
                dtpIssueDate.Value = parsedDate;
            }


            if (!string.IsNullOrEmpty(SelectionMemory.LogoPath) && File.Exists(SelectionMemory.LogoPath))
            {
                picLogo.Image = Image.FromFile(SelectionMemory.LogoPath);
            }

        }

        /// <summary>
        /// loads default project information from the Revit document
        /// </summary>
        private void LoadProjectInfoDefaults()
        {
            if (BOQData.Doc != null)
            {
                ProjectInfo projInfo = BOQData.Doc.ProjectInformation;

                txtProjectName.Text = string.IsNullOrWhiteSpace(SelectionMemory.ProjectName)
                    ? BOQData.Doc.Title ?? ""
                    : SelectionMemory.ProjectName;

                txtClientName.Text = string.IsNullOrWhiteSpace(SelectionMemory.ClientName)
                    ? projInfo.LookupParameter("Client Name")?.AsString() ?? ""
                    : SelectionMemory.ClientName;

                txtLocation.Text = string.IsNullOrWhiteSpace(SelectionMemory.Location)
                    ? projInfo.get_Parameter(BuiltInParameter.PROJECT_ADDRESS)?.AsString() ?? ""
                    : SelectionMemory.Location;

                string revitDate = projInfo.get_Parameter(BuiltInParameter.PROJECT_ISSUE_DATE)?.AsString();
                if (DateTime.TryParse(revitDate, out DateTime parsedDate))
                {
                    dtpIssueDate.Value = parsedDate;
                }
                else
                {
                    dtpIssueDate.Value = DateTime.Now;
                }
            }
            else
            {
                dtpIssueDate.Value = DateTime.Now;
            }
        }

        /// <summary>
        /// checks or unchecks a group of categories based on the provided BuiltInCategory array
        /// </summary>
        /// <param name="check"></param>
        /// <param name="categories"></param>
        private void ToggleGroup(bool check, BuiltInCategory[] categories)
        {
            for (int i = 0; i < clbCategories.Items.Count; i++)
            {
                var item = (CategoryItem)clbCategories.Items[i];
                if (categories.Contains(item.Bic))
                {
                    clbCategories.SetItemChecked(i, check);
                }
            }

        }


        // called when the user clicks the "Extract" button
        private void ExtractBtn_Click(object sender, EventArgs e)
        {
            try
            {
                var allowedCategories = GetSelectedCategories();
                BOQExtractor.Extract(BOQData.Doc, allowedCategories);

                lblCount.Text = $"Items extracted: {BOQData.Items.Count}";


                var groubedItems = BOQData.Items
                                   .GroupBy(it => new { it.Category, it.ItemCode, it.Description, it.Unit })
                                   .Select(g => new BOQItem
                                   (
                                   category: g.Key.Category,
                                   itemCode: g.Key.ItemCode,
                                   description: $"{g.Key.Category}:   {g.Key.Description}",
                                   g.Sum(x => x.Quantity), g.Key.Unit)
                                   {
                                       UsageCount = g.Count()
                                   }
                                   ).ToList();
                dgvItems.Rows.Clear();
                foreach(var item in groubedItems)
                {
                    dgvItems.Rows.Add(item.ItemCode,
                                      item.Description,
                                      item.Quantity + " " + item.Unit,
                                      item.UsageCount);
                }

                // Auto-size columns after filling
                dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during extraction: {ex.Message}",
                                                                            "Error",
                                                                            MessageBoxButtons.OK,
                                                                            MessageBoxIcon.Error);
            }
        }

        //called when the user clicks the "Export" button
        private void ExportBtn_Click(object sender, EventArgs e)
        {

            //exit if no item has been extracted
            if (dgvItems.Rows.Count == 0)
            {
                MessageBox.Show("Please extract items before exporting.",
                                "No Items Extracted",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return;
            }
            try
            {

                if (BOQData.Doc == null) 
                { 
                    MessageBox.Show("No active Revit document found. Please open a project first.",
                                    "Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                    return;
                }

                // Get project information
                ProjectInfo projInfo = BOQData.Doc.ProjectInformation;
                // Read parameters safely

                string projectName =string.IsNullOrWhiteSpace(txtProjectName.Text)?
                                    BOQData.Doc.Title ?? "N/A"
                                    :txtProjectName.Text;

                string clientName = string.IsNullOrWhiteSpace(txtClientName.Text) ?
                                    projInfo.LookupParameter("Client Name")?.AsString() ?? "N/A"
                                    :txtClientName.Text;

                string location = string.IsNullOrWhiteSpace(txtLocation.Text) ?
                                    projInfo.get_Parameter(BuiltInParameter.PROJECT_ADDRESS)?.AsString() ?? "N/A"
                                    :txtLocation.Text;

                string issueDate = dtpIssueDate.Value.ToString("dd/MM/yyyy");


                //asking the user to select the saving directory
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Title = "Save BOQ Export";
                    saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveDialog.FileName = $"{projectName}_BOQExport_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
                    
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Safety: ensure directory exists
                        string chosenPath = saveDialog.FileName;
                        string dir = Path.GetDirectoryName(chosenPath);
                        if (!Directory.Exists(dir))
                        { 
                            MessageBox.Show("Selected folder does not exist.",
                                            "Error",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error);
                            return; 
                        }
                        // set the saving directory
                        BOQData.SavingDir = new DirectoryInfo(dir);

                        //exporting the data to excel
                        BOQData.ReportFile = new FileInfo(chosenPath);
                        BOQData.ExportToExcel(projectName,
                                        clientName,
                                        location,
                                        issueDate
                                        );

                        MessageBox.Show($"BOQ Export completed successfully!\nSaved to:\n{chosenPath}",
                                        "Export",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);

                    }
                }                                    
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during export: {ex.Message}");
            }
        }

        // called when the user clicks the "Close" button
        private void CloseBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        // called when the user clicks the "Browse" button to select a logo image
        private void btnBrowseLogo_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDialog = new OpenFileDialog())
            {
                openDialog.Title = "Select Logo Image";
                openDialog.Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp";

                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    picLogo.Image = Image.FromFile(openDialog.FileName);
                    // logo is fit 
                    picLogo.SizeMode = PictureBoxSizeMode.Zoom;
                    SelectionMemory.LogoPath = openDialog.FileName;
                }
            }
        }

        // called when the user changes the selection in the category checklist box
        private void clbCategories_SelectedIndexChanged(object sender, EventArgs e)
        {

            // Delay update until after the check state changes
            this.BeginInvoke((MethodInvoker)delegate
            {
                UpdateSelectedItemsLabel();
            });
        }

        // this section handles the group checkboxes for different disciplines

        private void ArchCB_CheckedChanged(object sender, EventArgs e)
        {
            ToggleGroup(ArchCB.Checked, new[]
            {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_CurtainWallPanels,
                BuiltInCategory.OST_CurtainWallMullions,
                BuiltInCategory.OST_Ceilings,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_Casework,
                BuiltInCategory.OST_DetailComponents,
                BuiltInCategory.OST_Site,
                BuiltInCategory.OST_Topography,
                BuiltInCategory.OST_Parking
            });
            UpdateSelectedItemsLabel();

        }

        private void StruCB_CheckedChanged_1(object sender, EventArgs e)
        {
            ToggleGroup(StruCB.Checked, new[]
            {
                BuiltInCategory.OST_Columns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_Rebar,
                BuiltInCategory.OST_FabricAreas,
                BuiltInCategory.OST_FabricReinforcement,
                BuiltInCategory.OST_Walls // structural walls subset
            });
            UpdateSelectedItemsLabel();

        }

        private void MechCB_CheckedChanged(object sender, EventArgs e)
        {
            ToggleGroup(MechCB.Checked, new[]
            {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_Sprinklers,
                
            });
            UpdateSelectedItemsLabel();

        }

        private void ElecCB_CheckedChanged_1(object sender, EventArgs e)
        {
            ToggleGroup(ElecCB.Checked, new[]
           {
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_FireAlarmDevices,
                BuiltInCategory.OST_SecurityDevices,
                BuiltInCategory.OST_DataDevices,
                BuiltInCategory.OST_ElectricalCircuit,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting
            });
            UpdateSelectedItemsLabel();

        }



        //this section handles saving selections when the form is closing
        private void MainFrom_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Save categories
            SelectionMemory.SavedCategories = GetSelectedCategories();

            // Save group checkbox states
            SelectionMemory.ArchChecked = ArchCB.Checked;
            SelectionMemory.StruChecked = StruCB.Checked;
            SelectionMemory.MechChecked = MechCB.Checked;
            SelectionMemory.ElecChecked = ElecCB.Checked;

            // save meta data fields
            SelectionMemory.ProjectName = txtProjectName.Text;
            SelectionMemory.ClientName = txtClientName.Text;
            SelectionMemory.Location = txtLocation.Text;
            SelectionMemory.IssueDate = dtpIssueDate.Value.ToString("dd/MM/yyyy");

            // save logo path

            SelectionMemory.LogoPath = picLogo.Image != null ? SelectionMemory.LogoPath : "";


        }











        private void MainFrom_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        // change ui style from light to dark and vice versa
        private void metroToggle1_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void chkSiteworks_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void chkFireProtection_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}

