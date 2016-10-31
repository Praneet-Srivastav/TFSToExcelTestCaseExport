using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Excel = Microsoft.Office.Interop.Excel;

namespace TFSTestCaseExportToExcel
{
    public partial class FrmExport : Form
    {
        private TfsTeamProjectCollection _tfs;
        private ITestManagementTeamProject _teamProject;
        private ITestPlanCollection plans;
        private ITestSuiteEntryCollection testSuites;
        private ITestCaseCollection testCases;
        private ITestSuiteEntry suite;

        int flag1 = 0, flag2 = 0, flag3 = 0, flag4 = 0, flag5 = 0, flag6 = 0;

        public FrmExport()
        {
            InitializeComponent();
            //dtpCompleteStart.Value = DateTime.Today;
           // dtpCompleteEnd.Value = DateTime.Today;
            cmbRunType.SelectedIndex = 0;
        }

        private void btnTeamProject_Click(object sender, EventArgs e)
        {
            //Displaying the Team Project selection dialog to select the desired team project.
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.MultiProject, false);
            tpp.ShowDialog();

            //Following actions will be executed only if a team project is selected in the the opened dialog.
            if (tpp.SelectedTeamProjectCollection != null)
            {
                this._tfs = tpp.SelectedTeamProjectCollection;
                ITestManagementService test_service = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));
                this._teamProject = test_service.GetTeamProject(tpp.SelectedProjects[0].Name);

                //Populating the text field Team Project name (txtTeamProject) with the name of the selected team project.
                txtTeamProject.Text = tpp.SelectedProjects[0].Name;
                flag1 = 1;

                //Call to method "Get_TestPlans" to get the test plans in the selected team project
                Get_TestPlans(_teamProject);
            }

        }

        private void Get_TestPlans(ITestManagementTeamProject teamProject)
        {
            //Getting all the test plans in the collection "plans" using the query.
            this.plans = teamProject.TestPlans.Query("Select * From TestPlan");
            comBoxTestPlan.Items.Clear();
            flag2 = 0;
            comBoxTestSuite.Items.Clear();
            flag3 = 0;
            foreach (ITestPlan plan in plans)
            {
                //Populating the plan selection dropdown list with the name of Test Plans in the selected team project.
                comBoxTestPlan.Items.Add(plan.Name);
            }
        }

        private void Get_TestSuites(ITestSuiteEntryCollection Suites)
        {

            foreach (ITestSuiteEntry suite_entry in Suites)
            {
                this.suite = suite_entry;
                IStaticTestSuite newSuite = suite_entry.TestSuite as IStaticTestSuite;
                comBoxTestSuite.Items.Add(newSuite.Title);
            }
        }

        private void Get_TestCases(IStaticTestSuite testSuite)
        {
            this.testCases = testSuite.AllTestCases;
        }

        //Following method is invoked whenever a Test Plan is selected in the dropdown list.
        //Acording to the selected Test Plan in the dropdown list the Test Suites present in the selected Test Plan are populated in the Test Suite selection dropdown.
        private void comBoxTestPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            comBoxTestSuite.Items.Clear();
            int i = -1;
            if (comBoxTestPlan.SelectedIndex >= 0)
            {
                i = comBoxTestPlan.SelectedIndex;
                this.testSuites = plans[i].RootSuite.Entries;
                Get_TestSuites(testSuites);
                this.Cursor = Cursors.Arrow;
                flag2 = 1;
            }
        }

        private void btnFolderBrowse_Click(object sender, EventArgs e)
        {
            folderBrowserDialog.ShowDialog();
            txtSaveFolder.Text = folderBrowserDialog.SelectedPath;
            if (txtSaveFolder.Text != null || txtSaveFolder.Text != "")
            {
                flag4 = 1;
            }
        }

        private void comBoxTestSuite_SelectedIndexChanged(object sender, EventArgs e)
        {
            int j = -1;
            if (comBoxTestPlan.SelectedIndex >= 0)
            {
                j = comBoxTestSuite.SelectedIndex;
                this.suite = testSuites[j].TestSuite.TestSuiteEntry;
                IStaticTestSuite suite1 = suite.TestSuite as IStaticTestSuite;
                Get_TestCases(suite1);
                flag3 = 1;
                grpResultFilter.Enabled = true;
            }
        }

        private void FrmExport_Load(object sender, EventArgs e)
        {

        }

        //private void cbApplyDateFilter_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (cbApplyDateFilter.Checked)
        //    {
        //        dtpCompleteStart.Enabled = true;
        //        dtpCompleteEnd.Enabled = true;
        //    }
        //    else
        //    {
        //        dtpCompleteStart.Enabled = false;
        //        dtpCompleteEnd.Enabled = false;
        //    }
        //}

        private void lblWelcomeMessage_Click(object sender, EventArgs e)
        {

        }

        private void writeTestSteps(ref Excel.Worksheet xlWorkSheet, ITestStep tStep, ref int iRow, ref int iColumn)
        {
            string strStep = Regex.Replace(tStep.Title, @"<[^>]+>|&nbsp;", "").Trim();
            string strExpectedResult = Regex.Replace(tStep.ExpectedResult, @"<[^>]+>|&nbsp;", "").Trim();

            iColumn = 2;
            xlWorkSheet.Cells[iRow, 11] = strStep;
            xlWorkSheet.Cells[iRow, 12] = strExpectedResult;
            xlWorkSheet.Cells[iRow, 9] = "Test Step";
            iRow++;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //DateTime startDateTime = DateTime.Now;
            flag6 = 1;

            if (txtFileName.Text != null || txtFileName.Text != "")
            {
                flag5 = 1;
            }
            if (txtSaveFolder.Text != null || txtSaveFolder.Text != "")
            {
                flag4 = 1;
            }
            //if (cbApplyDateFilter.Checked && (DateTime.Compare(dtpCompleteEnd.Value, dtpCompleteStart.Value) >= 0))
            //{
            //    flag6 = 1;
            //}
            //else if (!cbApplyDateFilter.Checked)
            //{
            //    flag6 = 1;
            //}
            //else
            //{
            //    flag6 = 0;
            //}

            if (flag1 == 1 && flag2 == 1 && flag3 == 1 && flag4 == 1 && flag5 == 1 && flag6 == 1)
            {
                this.Cursor = Cursors.WaitCursor;
                btnExport.Enabled = false;
                btnCancel.Enabled = false;
                btnTeamProject.Enabled = false;
                btnFolderBrowse.Enabled = false;
                comBoxTestPlan.Enabled = false;
                comBoxTestSuite.Enabled = false;
                string filePath = string.Format(@"{0}\{1}.xlsx", txtSaveFolder.Text.Trim(), txtFileName.Text.Trim());

                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                Excel.Range chartRange;

                xlApp = new Excel.Application();
                xlWorkBook = null;

                if (!File.Exists(filePath))
                {
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkBook.SaveAs(filePath);
                }

                xlWorkBook = xlApp.Workbooks.Open(filePath);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                string wsName = (comBoxTestSuite.Text.Length > 15) ? comBoxTestSuite.Text.Substring(0, 15) : comBoxTestSuite.Text;
                xlWorkSheet.Name = string.Concat(xlWorkSheet.Name.Substring(5), string.Format(" {0}-{1}", wsName, cmbRunType.Text));

                xlWorkSheet.Cells[1, 1] = "ID";
                //xlWorkSheet.Cells[1, 2] = "Pass %";
               // xlWorkSheet.Cells[1, 3] = "Fail %";
               // xlWorkSheet.Cells[1, 4] = "RunCount";
               // xlWorkSheet.Cells[1, 5] = "LatestResult";
              //  xlWorkSheet.Cells[1, 6] = "Run Date";
               // xlWorkSheet.Cells[1, 7] = "State";
                xlWorkSheet.Cells[1, 8] = "Priority";
                xlWorkSheet.Cells[1, 9] = "WI Type";
                xlWorkSheet.Cells[1, 10] = "Title";
                xlWorkSheet.Cells[1, 11] = "Action";
                xlWorkSheet.Cells[1, 12] = "Expected Result";
              //  xlWorkSheet.Cells[1, 13] = "Complexity";
                xlWorkSheet.Cells[1, 14] = "Tags";
              //  xlWorkSheet.Cells[1, 15] = "Area Path";
              //  xlWorkSheet.Cells[1, 16] = "Iteration Path";
              //  xlWorkSheet.Cells[1, 17] = "Module/Feature";

                (xlWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 5;
                (xlWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 5;
                (xlWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 5;
                (xlWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 8;
                (xlWorkSheet.Columns["E", Type.Missing]).ColumnWidth = 10;
                (xlWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 12;
                (xlWorkSheet.Columns["G", Type.Missing]).ColumnWidth = 8;
                (xlWorkSheet.Columns["H", Type.Missing]).ColumnWidth = 6;
                (xlWorkSheet.Columns["I", Type.Missing]).ColumnWidth = 8;
                (xlWorkSheet.Columns["J", Type.Missing]).ColumnWidth = 40;
                (xlWorkSheet.Columns["K", Type.Missing]).ColumnWidth = 40;
                (xlWorkSheet.Columns["L", Type.Missing]).ColumnWidth = 40;
                (xlWorkSheet.Columns["M", Type.Missing]).ColumnWidth = 10;
                (xlWorkSheet.Columns["N", Type.Missing]).ColumnWidth = 30;
                (xlWorkSheet.Columns["O", Type.Missing]).ColumnWidth = 30;
                (xlWorkSheet.Columns["P", Type.Missing]).ColumnWidth = 30;
                (xlWorkSheet.Columns["Q", Type.Missing]).ColumnWidth = 20;


                int row = 2;
                int col = 1;
                string upperBound = "a";
                string lowerBound = "a";

                // cache tfs query results
                ITestRunHelper allTestRuns = null;
                IEnumerable<ITestRun> testRuns = null;
                ITestCaseResultHelper allTestCaseResults = null;

                try
                {
                    //Cache the testRuns
                    allTestRuns = _teamProject.TestRuns;

                    if (string.Compare(cmbRunType.Text.Trim(), "Manual") == 0)
                    {
                        testRuns = allTestRuns.Query("select TestRunID From TestRun where IsAutomated=0");
                    }
                    else if (string.Compare(cmbRunType.Text.Trim(), "Automated") == 0)
                    {
                        testRuns = allTestRuns.Query("select TestRunID From TestRun where IsAutomated=1");
                    }

                    //Cache testresults
                    allTestCaseResults = this._teamProject.TestResults;

                    foreach (ITestCase testCase in testCases)
                    {
                        upperBound = "a";
                        lowerBound = "a";
                        col = 1;

                        // Get the test results
                        var testCaseResult = allTestCaseResults.ByTestId(testCase.Id).Where(b => b.Outcome == TestOutcome.Passed || b.Outcome == TestOutcome.Failed);

                        // apply date filter to the results
                        //if (cbApplyDateFilter.Checked)
                        //{
                        //    testCaseResult = testCaseResult.Where(t => DateTime.Compare(t.DateCompleted.Date, dtpCompleteStart.Value.Date) >= 0 && DateTime.Compare(t.DateCompleted.Date, dtpCompleteEnd.Value.Date) <= 0);
                        //}

                        // apply Run Type filter
                        if (string.Compare(cmbRunType.Text.Trim(), "All") != 0)
                        {
                            testCaseResult = testCaseResult.Where(a => testRuns.FirstOrDefault(b => b.Id == a.TestRunId) != null);
                        }

                        // order the output in decending order to get the latest record
                        testCaseResult = testCaseResult.OrderByDescending(t => t.DateCompleted);

                        if (testCaseResult.Count() > 0)
                        {
                            int runCount = testCaseResult.Count();
                            float passCount = testCaseResult.Where(p => p.Outcome == TestOutcome.Passed).Count();
                            float failCount = testCaseResult.Where(p => p.Outcome == TestOutcome.Failed).Count();

                            float passPercent = ((passCount / runCount) * 100);
                            float failPercent = ((failCount / runCount) * 100);

                           // xlWorkSheet.Cells[row, 4] = runCount;
                            //xlWorkSheet.Cells[row, 2] = passPercent.ToString("00.00");
                           // xlWorkSheet.Cells[row, 3] = failPercent.ToString("00.00");
                           // xlWorkSheet.Cells[row, 5] = testCaseResult.FirstOrDefault().Outcome.ToString();
                           // xlWorkSheet.Cells[row, 6] = testCaseResult.FirstOrDefault().DateCompleted.ToString();
                        }
                        else
                        {
                            xlWorkSheet.Cells[row, 5] = "Not Run";
                        }

                        xlWorkSheet.Cells[row, 1] = testCase.Id.ToString();
                        //xlWorkSheet.Cells[row, 7] = testCase.State.ToString();
                        xlWorkSheet.Cells[row, 8] = testCase.Priority.ToString();
                        xlWorkSheet.Cells[row, 9] = "Test Title";
                        xlWorkSheet.Cells[row, 10] = testCase.Title.ToString();
                        //xlWorkSheet.Cells[row, 13] = testCase.WorkItem.Fields["Complexity"].OriginalValue.ToString();
                        xlWorkSheet.Cells[row, 14] = testCase.WorkItem.Tags.ToString();
                     //   xlWorkSheet.Cells[row, 15] = testCase.Area.ToString();
                      //  xlWorkSheet.Cells[row, 16] = testCase.WorkItem.IterationPath.ToString();
                      //  xlWorkSheet.Cells[row, 17] = testCase.WorkItem.NodeName.ToString();

                        upperBound += row;

                        row++;
                        TestActionCollection testActions = testCase.Actions;
                        foreach (ITestAction testAction in testActions)
                        {
                            var tstSharedStepRef = testAction as ISharedStepReference;
                            string strStep = string.Empty;
                            string strExpectedResult = string.Empty;

                            if (tstSharedStepRef != null)
                            {
                                ISharedStep tstSharedStep = tstSharedStepRef.FindSharedStep();
                                foreach (ITestAction tstSharedAction in tstSharedStep.Actions)
                                {
                                    ITestStep tstSharedTestStep = tstSharedAction as ITestStep;
                                    writeTestSteps(ref xlWorkSheet, tstSharedTestStep, ref row, ref col);
                                }
                            }
                            else {
                                ITestStep tstSharedTestStep = testAction as ITestStep;
                                writeTestSteps(ref xlWorkSheet, tstSharedTestStep, ref row, ref col);
                            }
                        }
                        lowerBound += (row - 1);
                    }
                    lowerBound = "q";
                    lowerBound += (row - 1);
                    chartRange = xlWorkSheet.get_Range("a1", "q1");
                    chartRange.Font.Bold = true;
                    chartRange.Interior.Color = 18018018;
                    xlWorkSheet.UsedRange.Replace("&gt;", ">");
                    xlWorkSheet.UsedRange.Font.Size = 9;
                    xlWorkSheet.UsedRange.WrapText = true;

                    MessageBox.Show(string.Format("Test Cases exported successfully to specified filepath: {0}", filePath));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    xlWorkBook.Save();
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet);
                    this.Cursor = Cursors.Arrow;

                    btnExport.Enabled = true;
                    btnCancel.Enabled = true;
                    btnTeamProject.Enabled = true;
                    btnFolderBrowse.Enabled = true;
                    comBoxTestPlan.Enabled = true;
                    comBoxTestSuite.Enabled = true;
                }
            }
            else
            {
                MessageBox.Show("All fields are not populated.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}