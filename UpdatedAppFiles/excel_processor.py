
import pandas as pd
import numpy as np
import openpyxl
import time
import random
from PyQt6.QtCore import QThread, pyqtSignal

from algorithmPolicy import (
    loadWorkloadPolicy, loadInstructorTrack, loadSpecialCourses,
    rowIsValid, Course, FacultyMember, adjust_co_convened
)

class ExcelProcessor(QThread):
    """Threaded Excel workload processor using updated algorithm."""
    progress = pyqtSignal(int)
    completed = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, raw_file_path, policy_file_path, track_file_path, special_file_path):
        super().__init__()
        self.raw_file_path = raw_file_path
        self.policy_file_path = policy_file_path
        self.track_file_path = track_file_path
        self.special_file_path = special_file_path

    def run(self):
        try:
            # 1) Load raw data
            converters = {
                'Max Units': lambda x: float(x) if pd.notna(x) else 0.0,
                'Enroll Total': lambda x: int(float(x)) if pd.notna(x) else 0
            }
            raw_df = pd.read_excel(self.raw_file_path, sheet_name='Raw Data', converters=converters)
            raw_df = raw_df[raw_df.apply(rowIsValid, axis=1)].reset_index(drop=True)
            raw_df = raw_df.drop_duplicates(subset=[
                'Instructor Emplid', 'Term', 'Subject', 'Cat Nbr', 'Section',
                'Start Date', 'Start Time', 'Facility Building', 'Facility Room'
            ])

            # 2) Supporting data
            policy = loadWorkloadPolicy(self.policy_file_path) if self.policy_file_path else loadWorkloadPolicy()
            tracks = loadInstructorTrack(self.track_file_path) if self.track_file_path else {}
            special = loadSpecialCourses(self.special_file_path) if self.special_file_path else set()

            # 3) Build structures
            faculty = {}
            courseGroups = {}
            other = {}

            for _, row in raw_df.iterrows():
                role = str(row.get('Instructor Role', '')).strip().upper()
                if role != 'PI':
                    other.setdefault(str(row.get('Instructor Emplid', '')), []).append(row.to_dict())
                    continue

                emplid_val = row.get('Instructor Emplid')
                if pd.isna(emplid_val):
                    continue
                emplid = int(float(emplid_val))
                if emplid not in tracks:
                    other.setdefault(str(emplid), []).append(row.to_dict())
                    continue

                course = Course(row.to_dict(), policy, special)
                key = course.getGroupKeyForGrouping()
                courseGroups.setdefault(key, []).append(course)

                if emplid not in faculty:
                    faculty[emplid] = FacultyMember(row.get('Instructor', ''), row.get('Instructor Email', ''), emplid, role, tracks[emplid])
                faculty[emplid].addCourse(course)

            # 4) Team‑taught division
            for lst in courseGroups.values():
                pi_only = [c for c in lst if str(c.instructorRole).strip().upper() == "PI"]
                unique_emplids = {c.instructorEmplid for c in pi_only}
                if len(unique_emplids) >= 2:
                    for c in pi_only:
                        c.adjustLoadDivision(len(unique_emplids))

            # 5) Co‑convened adjustment
            adjust_co_convened([c for lst in courseGroups.values() for c in lst], mode='collapse')

            # 6) Recalculate all loads after adjustments
            for fac in faculty.values():
                for c in fac.courses.values():
                    c.load = c.calculateLoad()

            # 7) Calculate summary
            summary_rows = []
            for fac in faculty.values():
                fac.calculateTotalLoad()
                units = sorted({getattr(c, 'unit', '') for c in fac.courses.values() if getattr(c, 'unit', '')})
                course_list = sorted({
                    f"{c.rawData.get('Subject', '').strip()} {c.catNbr}-{c.rawData.get('Section', '').strip()} - {c.rawData.get('Class Description', '').strip().title()}"
                    for c in fac.courses.values()
                })
                
                summary_rows.append({
                    'Instructor': fac.name,
                    'Emplid': fac.emplid,
                    'Track': fac.track or 'Unknown',
                    'Total Workload': round(fac.totalLoad, 2),
                    'Units Taught': ', '.join(units),
                    'Courses Taught': '; '.join(course_list)
                })

            summary_df = pd.DataFrame(summary_rows)

            # 8) Write output
            out_file = self.raw_file_path.replace('.xlsx', '_summary.xlsx')
            with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                raw_df.to_excel(writer, sheet_name='Processed Raw Data', index=False)
                summary_df.to_excel(writer, sheet_name='Faculty Summary', index=False)

            # Simulate progress
            pct = 0
            while pct < 100:
                pct = min(pct + random.randint(5, 10), 100)
                self.progress.emit(pct)
                time.sleep(0.03)

            self.completed.emit(out_file)

        except Exception as e:
            self.error.emit(str(e))
