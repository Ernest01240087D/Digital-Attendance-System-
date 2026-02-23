#include <iostream>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>

using namespace std;

/* =========================
   STUDENT CLASS
========================= */

class Student {
public:
    string name;
    string indexNumber;

    Student(string n, string i) {
        name = n;
        indexNumber = i;
    }

    void display() {
        cout << "Name: " << name
             << ", Index Number: " << indexNumber << endl;
    }
};

vector<Student> students;

                    //Attendance Session Class//
class AttendanceSession {
public:
    string courseCode;
    string date;
    string startTime;
    int duration;

    AttendanceSession(string c, string d, string s, int dur) {
        courseCode = c;
        date = d;
        startTime = s;
        duration = dur;
    }

    void displaySession() {
        cout << "Course: " << courseCode
             << ", Date: " << date
             << ", Start Time: " << startTime
             << ", Duration: " << duration << " hour(s)" << endl;
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
    string status; // Present, Absent, Late

    AttendanceRecord(string s, string c, string st) {
        studentIndex = s;
        courseCode = c;
        status = st;
    }
};

vector<AttendanceRecord> attendanceRecords;


/* =========================
   FUNCTIONS
========================= */
                    //Student registration function//
void registerStudent() {
    string name, index;

    cin.ignore();
    cout << "Enter Student Name: ";
    getline(cin, name);

    cout << "Enter Index Number: ";
    getline(cin, index);

    students.push_back(Student(name, index));

    // Save to file immediately
    ofstream file("students.txt", ios::app);
    file << name << "," << index << endl;
    file.close();

    cout << "\nStudent saved to students.txt:\n";
    cout << name << "," << index << endl;
}

void viewStudents() {
    if (students.empty()) {
        cout << "No students registered.\n";
        return;
    }

    for (auto &s : students)
        s.display();
}

void searchStudent() {
    string index;
    cout << "Enter index number to search: ";
    cin >> index;

    for (auto &s : students) {
        if (s.indexNumber == index) {
            s.display();
            return;
        }
    }
    cout << "Student not found.\n";
}
                      //Attendance session creation function//
void createSession() {
    string course, date, time;
    int duration;

    cout << "Enter Course Code: ";
    cin >> course;

    cout << "Enter Date (YYYY-MM-DD): ";
    cin >> date;

    cout << "Enter Start Time: ";
    cin >> time;

    cout << "Enter Duration (hours): ";
    cin >> duration;

    sessions.push_back(AttendanceSession(course, date, time, duration));
    // Save to file immediately
    ofstream file(" session_EE201_" + date + ".txt", ios::app);
    file << course << "," << date << "," << time << "," << duration << endl;
    file.close();

    cout << "\nSession saved to session_EE201_" + date + " .txt:\n";
    cout << course << "," << date << "," << time << "," << duration << endl;
}

                    //Attendance marking function//
void markAttendance() {
    if (students.empty() || sessions.empty()) {
        cout << "Students or sessions missing.\n";
        return;
    }

    cout << "\nAvailable Sessions:\n";
    for (int i = 0; i < sessions.size(); i++) {
        cout << i + 1 << ". ";
        sessions[i].displaySession();
    }

    int sessionChoice;
    cout << "Select session number: ";
    cin >> sessionChoice;

    if (sessionChoice < 1 || sessionChoice > sessions.size()) {
        cout << "Invalid session choice.\n";
        return;
    }

    AttendanceSession selectedSession = sessions[sessionChoice - 1];

    for (auto &s : students) {
        char response;
        cout << "Status for " << s.name << " (P=Present, A=Absent, L=Late): ";
        cin >> response;

        string status;
        if (response == 'P' || response == 'p') status = "Present";
        else if (response == 'L' || response == 'l') status = "Late";
        else status = "Absent";

        attendanceRecords.push_back(
            AttendanceRecord(s.indexNumber, selectedSession.courseCode, status)
        );
    }

    cout << "Attendance marked successfully!\n";
}

void updateAttendance() {
    if (attendanceRecords.empty()) {
        cout << "No attendance records to update.\n";
        return;
    }

    cout << "\nAttendance Records:\n";
    for (int i = 0; i < attendanceRecords.size(); i++) {
        cout << i + 1 << ". Index: " << attendanceRecords[i].studentIndex
             << ", Course: " << attendanceRecords[i].courseCode
             << ", Status: " << attendanceRecords[i].status << endl;
    }

    int choice;
    cout << "Select record number to update: ";
    cin >> choice;

    if (choice < 1 || choice > attendanceRecords.size()) {
        cout << "Invalid choice.\n";
        return;
    }

    char newStatus;
    cout << "Mark as (P=Present, A=Absent, L=Late): ";
    cin >> newStatus;

    if (newStatus == 'P' || newStatus == 'p') attendanceRecords[choice-1].status = "Present";
    else if (newStatus == 'L' || newStatus == 'l') attendanceRecords[choice-1].status = "Late";
    else attendanceRecords[choice-1].status = "Absent";

    cout << "Attendance updated successfully!\n";
}

/* =========================
   FUNCTIONS (REPORTS & FILES)
========================= */
void viewSessionAttendance() {
    string course;
    cout << "Enter course code: ";
    cin >> course;

    cout << "\nAttendance for " << course << ":\n";
    for (auto &r : attendanceRecords) {
        if (r.courseCode == course) {
            cout << "Index: " << r.studentIndex
                 << ", Status: " << r.status << endl;
        }
    }
}

void attendanceSummary() {
    int present=0, absent=0, late=0;
    for (auto &r : attendanceRecords) {
        if (r.status=="Present") present++;
        else if (r.status=="Late") late++;
        else absent++;
    }
    cout << "\nSummary:\n";
    cout << "Present: " << present << endl;
    cout << "Late: " << late << endl;
    cout << "Absent: " << absent << endl;
}

void loadStudents() {
    ifstream file("students.txt");
    string line;
    while (getline(file, line)) {
        stringstream ss(line);
        string name, index;
        getline(ss, name, ',');
        getline(ss, index);
        students.push_back(Student(name, index));
    }
    file.close();
}

void loadSessions() {
    ifstream file("session_EE201_YYYY_MM_DD.txt");
    string line;
    while (getline(file, line)) {
        stringstream ss(line);
        string c, d, t;
        int dur;
        getline(ss, c, ',');
        getline(ss, d, ',');
        getline(ss, t, ',');
        ss >> dur;
        sessions.push_back(AttendanceSession(c, d, t, dur));
    }
    file.close();
}


void saveAttendanceToFile() {
    ofstream file("attendance.txt");
    for (auto &r : attendanceRecords)
        file << r.studentIndex << "," << r.courseCode << "," << r.status << endl;
    file.close();
    cout << "Attendance saved successfully!\n";
}




/* =========================
   MAIN
========================= */
int main() {
    loadStudents();
    loadSessions();

    int choice;

    do {
        cout << "\n=====OPTIONS=====\n";
        cout << "1. Register Student\n";
        cout << "2. View Students\n";
        cout << "3. Search Student\n";
        cout << "4. Create Attendance Session\n";
        cout << "5. Mark Attendance\n";
        cout << "6. Update Attendance\n";
        cout << "7. View Session Attendance\n";
        cout << "8. Attendance Summary\n";
        cout << "9. Save Attendance to File\n";
        cout << "10. Exit\n";
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
            case 9: saveAttendanceToFile(); break;
            case 10: cout << "Exiting...\n"; break;
            default: cout << "Invalid choice!\n";
        }

    } while (choice != 10);


    return 0;
}
