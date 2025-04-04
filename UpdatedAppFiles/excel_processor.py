# excel_processor.py
# This is where you should put your caluclations @Cristan

import pandas as pd
import numpy as np
import openpyxl
import time
import random
from PyQt6.QtCore import QThread, pyqtSignal
from algorithmPolicy import (
    loadWorkloadPolicy,
    loadInstructorTrack,
    loadSpecialCourses,
    rowIsValid,
    Course,
    FacultyMember
)

class ExcelProcessor(QThread):
    progress = pyqtSignal(int)
    completed = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, file_paths, settings):
        super().__init__()
        self.file_paths = file_paths
        self.settings = settings
    
    
    def run(self):
        try:
            rawCourseData = pd.read_excel(self.file_paths["raw"])
            rawCourseData = rawCourseData[rawCourseData.apply(rowIsValid, axis=1)].reset_index(drop=True)
            policyParams = loadWorkloadPolicy(self.file_paths["policy"])
            instructorTracks = loadInstructorTrack(self.file_paths["instructorTrack"])
            specialCourses = loadSpecialCourses(self.file_paths["specialCourses"])

            facultyDict = {}
            otherStaff = {}
            courseGroups = {}

            for _, row in rawCourseData.iterrows():
               role = str(row.get('Instructor Role', '')).strip().upper()
                
               if role != "PI":
                  emplid = str(row.get('Instructor Emplid', '')).strip()
                  otherStaff.setdefault(emplid, []).append(row.to_dict())
                  continue

               emplid = int(float(row.get('Instructor Emplid', 0)))

               if emplid not in instructorTracks:
                  otherStaff.setdefault(str(emplid), []).append(row.to_dict())
                  continue

               name = row.get('Instructor', '')
               email = row.get('Instructor Email', '')
                
               courseObject = Course(row.to_dict(), policyParams, specialCourses)
               groupKey = courseObject.getGroupKey()
               courseGroups.setdefault(groupKey, []).append(courseObject)
                
               track = instructorTracks.get(emplid, None)
               if emplid not in facultyDict:
                  facultyDict[emplid] = FacultyMember(name, email, emplid, role, track)
               facultyDict[emplid].addCourse(courseObject)
               
            for groupKey, courses in courseGroups.items():
               if len(courses) > 1:
                  count = len(courses)
                  for courseObject in courses:
                     courseObject.adjustLoadDivision(count)
            
            for emplid, faculty in facultyDict.items():
                faculty.calculateTotalLoad()
                percentage = faculty.calculatePercentage()
            
            summary_data = []
            for emplid, faculty in facultyDict.items():
                percentage = faculty.calculatePercentage()
                summary_data.append({
                    "Faculty Name": faculty.name,
                    "Emplid": faculty.emplid,
                    "Total Workload Load (%)": round(faculty.totalLoad, 2),
                    "Workload Percentage": round(percentage, 2)
                })

            summary_data = []
            for emplid, faculty in facultyDict.items():
                percentage = faculty.calculatePercentage()
                units_set = set()
                for course in faculty.courses.values():
                    if hasattr(course, 'unit'):
                        units_set.add(course.unit)
                units_str = ", ".join(sorted(units_set))
                summary_data.append({
                    "Faculty Name": faculty.name,
                    "Emplid": faculty.emplid,
                    "Total Workload Load (%)": round(faculty.totalLoad, 2),
                    "Workload Percentage": round(percentage, 2),
                    "Units": units_str
                })
            summary_df = pd.DataFrame(summary_data)

            output_file = self.file_paths["raw"].replace(".xlsx", "_summary.xlsx")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                rawCourseData.to_excel(writer, sheet_name="Course Data", index=True)
                summary_df.to_excel(writer, sheet_name="Faculty Summary", index=False)
            
            progress = 0
            while progress < 100:
                increment = random.randint(1, 5)
                progress += increment
                progress = min(progress, 100)
                self.progress.emit(progress)
                time.sleep(random.uniform(0.1, 0.3))
            
            self.completed.emit(output_file)


        except Exception as e:
            self.error.emit(str(e))
