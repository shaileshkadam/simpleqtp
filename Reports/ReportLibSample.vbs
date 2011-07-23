'Declare an instance of the ReportLib class
Dim oReport
'New an instance of the ReportLib class
Set oReport = ReportLib

'Create a custom report with header info
oReport.CreateCustomReportFile "Header", "SGOWeb", "4.8.0", "1.0"

'Create a Test Suite Node
oReport.AddTestSuiteNode "Test Suite 1"

'Create a Test Case Node
oReport.AddTestCaseNode "TC_C_01", "Test Case 1"

'Create a passed step without expect/actual result
oReport.ReportPass array("Step 1"), false
'Create a passed step with Detail info
oReport.ReportPass array("Step 2","Detail info"), false

'Create a Test Case Node
oReport.AddTestCaseNode "TC_H_02", "Test Case 2"

'Create a failed step with screenshot
oReport.ReportFail array("Step 1", "Expected Result", "Actual Result"), true
oReport.ReportFail array("Step 2", "Expected Result", "Actual Result"), true

'Create a Test Suite Node
oReport.AddTestSuiteNode "Test Suite 2"

'Create a Test Case Node
oReport.AddTestCaseNode "TC_M_01", "Test Case 1"

'Create a passed step with linking to file
oReport.ReportPass array("Step 1", "Expected Result", "Actual Result", "_filepath\sample.txt"), false
oReport.ReportPass array("Step 2", "Expected Result", "Actual Result", "_filepath\sample.txt"), false

'Create a Test Case Node
oReport.AddTestCaseNode "TC_L_02", "Test Case 2"
'Create a failed step with screenshot and linking to file
oReport.ReportFail array("Step 1", "Expected Result", "Actual Result", "_filepath\sample.txt"), true
oReport.ReportFail array("Step 2", "Expected Result", "Actual Result", "_filepath\sample.txt"), true

'Release the instance
Set oReport = nothing
