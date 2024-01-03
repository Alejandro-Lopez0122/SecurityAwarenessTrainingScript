# This was last updated on 1/03/2024
$userName = "Alejandro Lopez" # Change this variable to the name of the user of your machine.

# Define the base folder path for the Security Awareness Training Status (FINAL) directory
$baseFolder = "C:\Users\$userName\Documents\Security Awareness Training Status (FINAL)"

# Create a hashtable of key-value pairs where each key is a name, and each value is a file path
$paths = @{
    "PathSATStatusReport" = Join-Path "C:\Users\$userName\Documents" "SAT Status Report"
    "PathSecurityAwarenessTrainingStatus" = $baseFolder
    "PathAcademicAffairsAdministration" = Join-Path $baseFolder "Academic Affairs\Academics Affairs Administration"
    "PathAdministrationAndFinance" = Join-Path $baseFolder "Administration and Finance"
    "PathAuxiliary" = Join-Path $baseFolder "Auxiliary"
    "PathCollegeOfArts" = Join-Path $baseFolder "College of Arts, Media & Communication"
    "PathCollegeOfEducation" = Join-Path $baseFolder "College of Education"
    "PathCollegeOfEngCompSci" = Join-Path $baseFolder "College of Eng & Comp Sci"
    "PathCollegeOfHealthHumanDev" = Join-Path $baseFolder "College of Health & Human Dev"
    "PathCollegeOfHumanities" = Join-Path $baseFolder "College of Humanities"
    "PathCollegeOfScienceMath" = Join-Path $baseFolder "College of Science and Math"
    "PathCollegeOfSocialBehaviorSci" = Join-Path $baseFolder "College of Social & Behavior Sci"
    "PathDavidNazarian" = Join-Path $baseFolder "David Nazarian Col of Bus&Econ"
    "PathInformationTechnology" = Join-Path $baseFolder "Information Technology"
    "PathStudentAffairs" = Join-Path $baseFolder "Student Affairs"
    "PathTsengCollege" = Join-Path $baseFolder "Tseng College"
    "PathUniversityAdvancement" = Join-Path $baseFolder "University Advancement"
    "PathUniversityLibrary" = Join-Path $baseFolder "University Library"
}

# Iterate through each key in the hashtable to insure the path exists
# If the path (Folder) does not exist, it will create the folder
foreach ($key in $paths.Keys) {
    if (!(Test-Path -Path $paths[$key])) {
        mkdir $paths[$key]
    }
}

# Create another hashtable where keys represent departments/sections and values represent subdirectory paths
$CollegePaths = @{
    "Academic Affairs" = "Academic Affairs\Academics Affairs Administration"
    "ED_OPP_PRG" = "Academic Affairs\Academics Affairs Administration"
    "Faculty Affairs" = "Academic Affairs\Academics Affairs Administration"
    "Undergraduate Studies" = "Academic Affairs\Academics Affairs Administration"
    "Graduate Studies" = "Academic Affairs\Academics Affairs Administration"
    "Institutional Research" = "Academic Affairs\Academics Affairs Administration"
    "Academic Resources" = "Academic Affairs\Academics Affairs Administration"
    "Provosts Office" = "Academic Affairs\Academics Affairs Administration"
    "Faculty Senate" = "Academic Affairs\Academics Affairs Administration"
    "Athletics" = "Administration and Finance"
    "Budget Planning" = "Administration and Finance"
    "Facilities Planning" = "Administration and Finance"
    "Financial Services" = "Administration and Finance"
    "Human Resources" = "Administration and Finance"
    "Internal Audit" = "Administration and Finance"
    "Police Services" = "Administration and Finance"
    "VPAC" = "Administration and Finance"
    "Administration" = "Administration and Finance"
    "AS" = "Auxiliary"
    "President's Office" = "Auxiliary"
    "TUC" = "Auxiliary"
    "USU" = "Auxiliary"
    "College of Arts" = "College of Arts, Media & Communication"
    "College of Education" = "College of Education"
    "College of Eng & Comp Sci" = "College of Eng & Comp Sci"
    "College of Health & Human Dev" = "College of Health & Human Dev"
    "College of Humanities" = "College of Humanities"
    "College of Science and Math" = "College of Science and Math"
    "College of Social & Behavior Sci" = "College of Social & Behavior Sci"
    "David Nazarian Col of Bus&Econ" = "David Nazarian Col of Bus&Econ"
    "Academic Technology" = "Information Technology"
    "Infrastructure Services" = "Information Technology"
    "IT Administration and Support Services" = "Information Technology"
    "Information Services" = "Information Technology"
    "Career Center" = "Student Affairs"
    "Center on Deafnesss" = "Student Affairs"
    "Counseling Services" = "Student Affairs"
    "Disability Resources" = "Student Affairs"
    "National Center on Deafness" = "Student Affairs"
    "Residence Life" = "Student Affairs"
    "Student Affairs Technology" = "Student Affairs"
    "Student Affairs" = "Student Affairs"
    "Student Health Center" = "Student Affairs"
    "Student Involvement" = "Student Affairs"
    "Student Affairs VP Office" = "Student Affairs"
    "Tseng College" = "Tseng College"
    "Public Relations" = "University Advancement"
    "University Advancement" = "University Advancement"
    "University Development" = "University Advancement"
    "Alumni Relations" = "University Advancement"
    "KCSN Radio Station" = "University Advancement"
    "University Library" = "University Library"
}

Get-ChildItem -Path $baseFolder -Include *.* -File -Recurse | Remove-Item -ErrorAction Ignore # Deletes all files within this folder structure. Keeps folder structure intact.

Get-ChildItem -Path $paths["PathSATStatusReport"] -File | Remove-Item -ErrorAction Ignore # Deletes all files in the SAT Status Report directory

$rawFileName = "ListDIV_Full Data_data.xlsx"
$SATStatusReportPath = $paths["PathSATStatusReport"]
$rawFilePath = Join-Path $SATStatusReportPath $rawFileName # Sets a variables for the raw file path
$sortedFilePath = Join-Path $SATStatusReportPath "SortedSAT.xlsx" # Sets a variable for the sorted file path

# This IF Statement checks to see if the file exists. If it does, it deletes it.
if (Test-Path -Path $rawFilePath) { Remove-Item -Path $rawFilePath -ErrorAction Ignore } # Deletes the old Raw workbook in folder 

# This IF Statement checks to see if the file exists. If it does, it deletes it.
if (Test-Path -Path $sortedFilePath) { Remove-Item -Path $sortedFilePath -ErrorAction Ignore } # Deletes the SortedSAT workbook in folder 

$objExcel = New-Object -ComObject Excel.Application # Creates Excel Object
$objExcel.Visible = $false # Enables/Disbaled whether you want to see the GUI or not

Copy (Join-Path "C:\Users\$userName\Downloads" $rawFileName) $SATStatusReportPath # Copies the raw excel workbook from your downloads and pastes it into "C:\Users$userName\Documents\SAT Status Report"

# The following function does all the formating to the raw file and saves it to a new file called SortedSAT.xlsx
function sortFullData($rawFilePath) {
    $Workbook = $objExcel.Workbooks.Open($rawFilePath)
    $OldWorksheet = $Workbook.Sheets.Item(1)

    $desiredHeaders = "Full Name", "College/Area", "Department", "Dept Id",
                      "Email Address", "Type", "Confidential", "CSUN ID",
                      "Division", "Phone", "Completion Dt", "Hire Dt", "Days since Hire", "Over Due"

    # Create a new worksheet
    $NewWorksheet = $Workbook.Sheets.Add()
    $NewWorksheet.Name = "NewSheet"

    # Loop through desired headers
    for ($i = 0; $i -lt $desiredHeaders.Length; $i++) {
        # Find the matching column in the old worksheet
        $totalColumns = $OldWorksheet.UsedRange.Columns.Count
        for ($j = 1; $j -le $totalColumns; $j++) {
            if ($OldWorksheet.Cells.Item(1, $j).Text -eq $desiredHeaders[$i]) {
                # Copy the column to the new worksheet
                $OldWorksheet.Columns.Item($j).EntireColumn.Copy() | Out-Null
                $NewWorksheet.Columns.Item($i + 1).EntireColumn.PasteSpecial() | Out-Null
                break
            }
        }
    }

    # Delete the old worksheet
    $OldWorksheet.Delete()

    # Sorts all Columns by College/Area
    $objRange = $NewWorksheet.UsedRange 
    $objRange2 = $NewWorksheet.Range("B1")  
    [void] $objRange2.Sort($objRange2,1,$null,$null,1,$null,1,1) 

    $NewWorksheet.Cells.Item(1,4).value() = "Department Number" # Renames "Dept Id" -> Department Number
    $NewWorksheet.Cells.Item(1,11).value() = "Last Completion Date" # Renames "Completion Dt" -> Last Completion Date
    $NewWorksheet.Cells.Item(1,12).value() = "Hire Date" # Renames "Hire Dt" -> Hire Date
    $NewWorksheet.Cells.Item(1,13).value() = "Days Since Hire" # Renames "Days since Hire" -> Days Since Hire
    $NewWorksheet.Cells.Item(1,14).value() = "Days Over Due" # Renames "Over Due" -> Days Over Due

    $NewWorksheet.Columns.item('H').NumberFormat = "000000000" # Formats CSUN ID # -> 000000000 (9 digits, forces leading zeros)

    # Loops through "Last Completion Date" column, replaces "1/1/1111" with "SAT Never Completed"
    $lastCompletionColumn = 11 # Set this to your "Last Completion Date" column number
    $totalRows = $NewWorksheet.UsedRange.Rows.Count
    for ($i = 2; $i -le $totalRows; $i++) {
        if ($NewWorksheet.Cells.Item($i, $lastCompletionColumn).Text -eq "1/1/1111") {
            $NewWorksheet.Cells.Item($i, $lastCompletionColumn).Value2 = "SAT Never Completed"
        }
    }

    $Workbook.SaveAs("C:\Users\$userName\Documents\SAT Status Report\SortedSAT.xlsx") # Saves the workbook to defined location with its new name defined by the CollegeArea variable
    $Workbook.close($true) # Closes the new workbook

    Write-Host "`nFile has been sorted"
}

function createCollegeArray($sortedFilePath) {
    $Workbook = $objExcel.Workbooks.Open($sortedFilePath) # Sets the workbook to the sortedSATList
    $Worksheet = $Workbook.Sheets.Item("NewSheet") # Sets the worksheet within the workbook
    
    $numOfRows = $Worksheet.UsedRange.Rows.Count # Counts the total amount of rows in the Workbook

    $collegeHashSet = New-Object System.Collections.Generic.HashSet[string] # Defines a HashSet for college names

    # Iterates through College/Area column and adds unique college/areas to the collegeHashSet
    for($i = 2; $i -le $numOfRows; $i++) {
        $collegeName = $Worksheet.cells.Item($i,2).text
        $collegeHashSet.Add($collegeName) | Out-Null
    }

    $Workbook.close($true) # Closes the new workbook

    # Converts the HashSet back to an array for returning
    $collegeArray = @($collegeHashSet)

    return $collegeArray
}

function createCollegeWorkbooks($Workbook, $CollegeArea, $sortedFilePath) {
    $Worksheet = $Workbook.Sheets.Item("NewSheet") # Sets the worksheet within the workbook
    
    $numOfRows = $Worksheet.UsedRange.Rows.Count # Counts the total amount of rows in the Workbook

    $getName = $Worksheet.Range("B1:B$numOfRows").find("$CollegeArea")
    $startRow = $getName.row 
    $count = 0

    for($i = $startRow; $i -lt $numOfRows + 1; $i++) {
        if ($Worksheet.cells.Item($i,2).text -eq $CollegeArea) {
            $count++ # Count the number of times that particular College/Area appears
        }
    }

    if($CollegeArea -eq "College of Eng/Comp Sci") { # Renames the "College of Eng/Comp Sci" file
        $CollegeArea = "College of Eng & Comp Sci"
    }

    if($CollegeArea -eq "College of Social/Behavior Sci") { # Renames the "College of Social/Behavior Sci" file
        $CollegeArea = "College of Social & Behavior Sci"
    }
    
    $lastRow = $startRow + $count - 1 # Gets the row number of the last occurence of that particular College/Area

    $range = $Worksheet.Range(“A$startRow : N$lastRow”) # Defines the range to copy
    $range.Copy() | Out-Null # Copy the range defined in the line above

    $newWorkBook = $objExcel.Workbooks.add() # Opens a completely new workbook
    $newWorkSheet = $newWorkBook.worksheets.Item(1) # Opens a new worksheet in the newly opened workbook

    $adjustedRange = $count + 3 # Makes our paste range 3 rows lower to fit in description, date, and headers
    $range = $newWorkSheet.Range("A4 : N$adjustedRange") # Defines the range to copy to
    $newWorkSheet.Paste($range) # Pastes content

    # Header Portion
    $header = $Worksheet.Range(“A1 : N1”) # Defines the range to copy for the header
    $header.Copy() | Out-Null # Copy the range defined in the line above
    $header = $newWorkSheet.Range("A3 : N3") # Defines the range to copy to for the header
    $newWorkSheet.Paste($header) # Pastes content

    # Sets Description and Date in the Spreadsheet
    $date = Get-Date -UFormat "%m/%d/%Y" # Sets a variable for the current Date
    $newWorkBook.worksheets.item(1).cells.item(1,1) = "Employees with Overdue Security Awareness Training" # Inserts the text into cell(1,1)
    $newWorkBook.worksheets.item(1).cells.item(2,1) = "Date Created: $date" # Inserts current Date into cell (2,1)

    # Get the college path from the dictionary, or default to the college name
    $collegePath = $CollegePaths[$CollegeArea]
    if($collegePath -eq $null) {
        $collegePath = $CollegeArea
    }

    $newWorkBook.SaveAs("C:\Users\$userName\Documents\Security Awareness Training Status (FINAL)\$collegePath\$CollegeArea.xlsx") 

    $newWorkBook.close($true) # Closes the new workbook
}

sortFullData $rawFilePath

$collegeArray = createCollegeArray $sortedFilePath
$totalColleges = $collegeArray.Length

# Open the workbook outside of the function
$Workbook = $objExcel.Workbooks.Open($sortedFilePath)

$i = 1

# Create the workbooks for each college
foreach ($college in $collegeArray) {
    createCollegeWorkbooks $Workbook $college $sortedFilePath    
    Write-Host "Workbook $i complete out of $totalColleges $college"
    $i = $i + 1   
}

# Close the workbook after creating all workbooks
$Workbook.Close($true)