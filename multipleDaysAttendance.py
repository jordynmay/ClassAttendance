# Example call:
# python3 multipleDaysAttendance.py Attendance.xlsx "101(MWF)"
import attendance
import os
import sys

def main() -> None:
    # Directory of .csv Zoom report files
    directory = "zoom"

    # .xlsx file of attendance records
    attendanceFile = sys.argv[1]

    # Section number, either 101(MWF) or 102(TTH)
    sectionNo = sys.argv[2]

    nameList = []

    #datesList = ["2021-01-20", "2021-01-22",
    #"2021-01-25", "2021-01-27"]

    ctr = 0
    for filename in os.listdir(directory):
        if(filename.endswith(".csv")):
            #print(filename[:-4])

            # Get location of report file
            zoomReport = directory+"/"+filename

            # Parse zoomReport to get a list of names of students who attended
            nameList = attendance.readZoomReport(zoomReport)

            #dateTime = attendance.getDateTime(datesList[ctr])

            # Assumes the file name is YYYY-MM-DD
            dateTime = attendance.getDateTime(filename[:-4])

            # Update the attendance for all students on date dateTime
            attendance.updateGrades(attendanceFile, sectionNo, nameList, dateTime)
        ctr += 1


if __name__ == "__main__":
    main()