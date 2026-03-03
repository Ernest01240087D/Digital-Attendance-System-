#include <iostream>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>

using namespace std;

/* =========================
   STUDENT 
========================= */
class Student {
public:
    string name;
    string indexNumber;

    Student(string n, string i) : name(n), indexNumber(i) {}

    void display() {
        cout << "Name: " << name << ", Index Number: " << indexNumber << endl;
    }
};

vector<Student> students;

/* =========================
   ATTENDANCE SESSION 
========================= */
class AttendanceSession {
public:
    string courseCode;
    string date;
    string startTime;
    int duration;

    AttendanceSession(string c, string d, string s, int dur) 
        : courseCode(c), date(d), startTime(s), duration(dur) {}

    void displaySession() {
        cout << "Course: " << courseCode << ", Date: " << date
             << ", Start Time: " << startTime << ", Duration: " << duration << " hour(s)" << endl;
    }
};

vector<AttendanceSession> sessions;

/* =========================
   ATTENDANCE RECORD
========================= */
class AttendanceRecord {
public:
    string studentIndex;
    string courseCode;
    string status; 

    AttendanceRecord(string s, string c, string st) 
        : studentIndex(s), courseCode(c), status(st) {}
};

vector<AttendanceRecord> attendanceRecords;

/* =========================
   NEW EXCEL EXPORT FUNCTION
========================= */
void exportToExcel() {
    if (attendanceRecords.empty()) {
        cout << "No records available to export.\n";
        return;
    }

    // Creating a .csv file which Excel opens perfectly
    ofstream file("Attendance_Report.csv");

    // Writing the Header Row
    file << "Student Index,Course Code,Status" << endl;

    // Writing the Data
    for (auto &r : attendanceRecords) {
        file << r.studentIndex << "," << r.courseCode << "," << r.status << endl;
    }

    file.close();
    cout << "Success! Data exported to 'Attendance_Report.csv'. Open this file in Excel.\n";
}

/* =========================
   EXISTING FUNCTIONS
========================= */
void registerStudent() {
    string name, index;
    cin.ignore();
    cout << "Enter Student Name: ";
    getline(cin, name);
    cout << "Enter Index Number: ";
    getline(cin, index);

    students.push_back(Student(name, index));
    ofstream file("students.txt", ios::app);
    file << name << "," << index << endl;
    file.close();
}

void viewStudents() {
    if (students.empty()) { cout << "No students registered.\n"; return; }
    for (auto &s : students) s.display();
}

void searchStudent() {
    string index;
    cout << "Enter index number to search: ";
    cin >> index;
    for (auto &s : students) {
        if (s.indexNumber == index) { s.display(); return; }
    }
    cout << "Student not found.\n";
}

void createSession() {
    string course, date, time;
    int duration;
    cout << "Enter Course Code: "; cin >> course;
    cout << "Enter Date (YYYY-MM-DD): "; cin >> date;
    cout << "Enter Start Time: "; cin >> time;
    cout << "Enter Duration (hours): "; cin >> duration;

    sessions.push_back(AttendanceSession(course, date, time, duration));
    ofstream file("session_" + course + "_" + date + ".txt", ios::app);
    file << course << "," << date << "," << time << "," << duration << endl;
    file.close();
}

void markAttendance() {
    if (students.empty() || sessions.empty()) {
        cout << "Students or sessions missing.\n";
        return;
    }
    cout << "\nAvailable Sessions:\n";
    for (int i = 0; i < (int)sessions.size(); i++) {
        cout << i + 1 << ". ";
        sessions[i].displaySession();
    }
    int choice;
    cout << "Select session number: "; cin >> choice;

    if (choice < 1 || choice > (int)sessions.size()) return;

    AttendanceSession selected = sessions[choice - 1];
    for (auto &s : students) {
        char res;
        cout << "Status for " << s.name << " (P/A/L): "; cin >> res;
        string stat = (res == 'P' || res == 'p') ? "Present" : (res == 'L' || res == 'l' ? "Late" : "Absent");
        attendanceRecords.push_back(AttendanceRecord(s.indexNumber, selected.courseCode, stat));
    }
}

void updateAttendance() {
    if (attendanceRecords.empty()) return;
    int choice;
    cout << "Select record (1-" << attendanceRecords.size() << "): "; cin >> choice;
    if (choice < 1 || choice > (int)attendanceRecords.size()) return;

    char newStat;
    cout << "New status (P/A/L): "; cin >> newStat;
    attendanceRecords[choice-1].status = (newStat == 'P' || newStat == 'p') ? "Present" : (newStat == 'L' || newStat == 'l' ? "Late" : "Absent");
}

void viewSessionAttendance() {
    string course;
    cout << "Enter course code: "; cin >> course;
    for (auto &r : attendanceRecords) {
        if (r.courseCode == course) cout << r.studentIndex << ": " << r.status << endl;
    }
}

void attendanceSummary() {
    int p=0, a=0, l=0;
    for (auto &r : attendanceRecords) {
        if (r.status == "Present") p++; else if (r.status == "Late") l++; else a++;
    }
    cout << "\nP: " << p << " L: " << l << " A: " << a << endl;
}

void loadStudents() {
    ifstream file("students.txt");
    string line;
    while (getline(file, line)) {
        stringstream ss(line);
        string n, i;
        getline(ss, n, ',');
        getline(ss, i);
        if(!n.empty()) students.push_back(Student(n, i));
    }
}

/* =========================
   MAIN
========================= */
int main() {
    loadStudents();
    int choice;
    do {
        cout << "\n--- ATTENDANCE SYSTEM ---\n";
        cout << "1. Register Student\n2. View Students\n3. Search Student\n";
        cout << "4. Create Session\n5. Mark Attendance\n6. Update Attendance\n";
        cout << "7. View Attendance\n8. Summary\n9. Export to Excel (CSV)\n10. Exit\n";
        cout << "Choice: ";
        cin >> choice;

        switch (choice) {
            case 1: registerStudent(); break;
            case 2: viewStudents(); break;
            case 3: searchStudent(); break;
            case 4: createSession(); break;
            case 5: markAttendance(); break;
            case 6: updateAttendance(); break;
            case 7: viewSessionAttendance(); break;
            case 8: attendanceSummary(); break;
            case 9: exportToExcel(); break;
            case 10: cout << "Goodbye!\n"; break;
        }
    } while (choice != 10);
    return 0;
}
