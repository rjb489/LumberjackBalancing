import pandas as pd
from collections import defaultdict
from typing import Iterable, Dict, List, Tuple, Set

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def _norm(s):
    return str(s).strip().lower() if pd.notna(s) else ""


def _meeting_signature(row_or_course) -> Tuple:
    data = row_or_course.rawData if hasattr(row_or_course, "rawData") else row_or_course
    return (
        data.get("Start Date"),
        data.get("End Date"),
        data.get("Start Time"),
        data.get("End Date"),
        data.get("Days"),
        #_norm(data.get("Facility Building")),
        #_norm(data.get("Facility Room")),
    )

# ---------------------------------------------------------------------------
# Policy / loaders
# ---------------------------------------------------------------------------

def loadWorkloadPolicy(path: str | None = None) -> dict:
    defaults = {
        "supplementalInstructionRate": 1.0,
        "specialCoursesRate": 0.005,
        "independentStudyRateHigh": 0.5,
        "independentStudyRateLow": 0.25,
        "laboratoryRate": 5.0,
        "lectureRate": 3.33,
        "maxLoadCap": 5.0,
        "lectureThreshold": {"low": 90, "mid": 150, "high": 200},
        "midRate": 4.17,
        "highRate": 5.0,
        "maxRate": 6.66
    }
    policy = defaults.copy()
    if path:
        try:
            df = pd.read_excel(path)
            kv = {str(k).strip(): v for k, v in zip(df.iloc[:, 0], df.iloc[:, 1])}
            if all(k in kv for k in ("lectureThreshold_low", "lectureThreshold_mid", "lectureThreshold_high")):
                policy["lectureThreshold"] = {
                    "low": float(kv.pop("lectureThreshold_low")),
                    "mid": float(kv.pop("lectureThreshold_mid")),
                    "high": float(kv.pop("lectureThreshold_high")),
                }
            for k, v in kv.items():
                try:
                    policy[k] = float(v)
                except (TypeError, ValueError):
                    policy[k] = v
        except Exception as e:
            print("Warning: failed to load policy – using defaults:", e)
    return policy


def loadInstructorTrack(p: str) -> dict:
    try:
        df = pd.read_excel(p)
        return pd.Series(df["Track"].values, index=df["Instructor Emplid"]).to_dict()
    except Exception as e:
        print("Warning loading track file:", e)
        return {}


def loadSpecialCourses(p: str) -> set:
    try:
        df = pd.read_excel(p)
        return set(df["Course"].dropna().astype(str).str.strip().str.lower())
    except Exception as e:
        print("Warning loading special courses:", e)
        return set()

# ---------------------------------------------------------------------------
# Row filter
# ---------------------------------------------------------------------------

def rowIsValid(row: pd.Series) -> bool:
    required = ["Instructor Role", "Instructor Emplid"]
    if any(pd.isna(row.get(c)) for c in required):
        return False
    
    return True

# ---------------------------------------------------------------------------
# Course object
# ---------------------------------------------------------------------------

class Course:
    def __init__(self, data: dict, policy: dict, special: Set[str]):
        self.rawData = data
        self.policy = policy
        self.special = special

        self.courseCategory = _norm(data.get("Course Category (CCAT)"))
        self.classCat = _norm(data.get("Class"))

        classNbrRaw = str(data.get("Class Nbr", "")).strip()
        self.classNbr = str(int(float(classNbrRaw))) if classNbrRaw.replace(".", "", 1).isdigit() else classNbrRaw

        cat_raw = str(data.get("Cat Nbr", "")).strip()
        self.catNbr = str(int(float(cat_raw))) if cat_raw.replace(".", "", 1).isdigit() else cat_raw
        self.instructorRole = _norm(data.get("Instructor Role"))

        numUnits = data.get("Max Units", 0)
        self.maxUnits = float(numUnits) if pd.notna(numUnits) else 0.0

        numEnroll = data.get("Enroll Total", 0)
        self.enrollTotal = int(numEnroll) if pd.notna(numEnroll) else 0

        emplid_raw = data.get("Instructor Emplid", None)
        self.instructorEmplid = int(float(emplid_raw)) if pd.notna(emplid_raw) else None

        self.startDate = data.get("Start Date")
        self.startTime = data.get("Start Time")
        self.facilityBuilding = data.get("Facility Building")
        self.facilityRoom = data.get("Facility Room")

        self.unit = str(data.get("Unit", "")).strip()
        self.load: float | None = None

        self.co_convened_members: List[str] = []
        self.team_taught_members: List[str] = []

    # ------------------------------------------------------------------
    def _meeting_signature(self):
        return _meeting_signature(self.rawData)

    def getGroupKeyForGrouping(self):
        term = self.rawData.get("Term")
        subject = self.rawData.get("Subject")
        section = self.rawData.get("Section")
        categoryNbr = self.rawData.get("Cat Nbr")
        classNbr = self.rawData.get("Class Nbr")
        classCat = self.rawData.get("Class")
        return (term, subject, categoryNbr, section, classNbr, classCat) + self._meeting_signature()

    def getGroupKeyForCollapsing(self):
        term = self.rawData.get("Term")
        subject = self.rawData.get("Subject")
        section = self.rawData.get("Section")
        return (self.instructorEmplid, term, subject, section) + self._meeting_signature()

    # ------------------------------------------------------------------
    def _baseRate(self):
        p = self.policy
        if any(k in self.courseCategory for k in ("independent study", "research", "fieldwork", "research - experiential", "individualized study - experie")):
            if self.maxUnits > 0 and self.maxUnits <= 2:
                return float(p.get("independentStudyRateLow", 0.25))
            elif self.maxUnits > 2:
                return float(p.get("independentStudyRateHigh", 0.5))
        if "laboratory" in self.courseCategory:
            return float(p.get("laboratoryRate", 5.0))
        if any(k in self.classCat for k in ("mat 100", "mat 108", "mat 114", "mat 125")) and self.instructorRole == "st":
            return float(p.get("supplementalInstructionRate", 1.0))
        if any(k in self.catNbr for k in ("699", "799")):
            return 1.0
        return float(p.get("lectureRate", 3.33))

    def _adjustForEnrollment(self, base):
        if "lecture" not in self.courseCategory:
            return base
        p = self.policy
        low, mid, high = (float(p["lectureThreshold"][k]) for k in ("low", "mid", "high"))
        s = self.enrollTotal
        if s < low:
            return base
        if low <= s <= mid:
            return float(p["midRate"])
        if mid < s <= high:
            return float(p["highRate"])
        
        return float(p["maxRate"])
        

    # ------------------------------------------------------------------
    def calculateLoad(self):
        if self.enrollTotal == 0 or self.enrollTotal is None:
            self.load = 0.0
            return 0.0
        
        if self.maxUnits == 0 or self.maxUnits is None:
            self.load = 0.0
            return 0.0
        
        if self.load is not None:
            return self.load
        
        base = self._baseRate()
        eff_enroll = self.enrollTotal

        if any(k in self.courseCategory for k in ("independent study", "research", "fieldwork")):
            load = base * eff_enroll
            cap = self.policy.get("maxLoadCap", 5.0)
            load = min(load, cap)

        elif any(k in self.classCat for k in ("mat 100", "mat 108", "mat 114", "mat 125")) and self.instructorRole == "st":
            load = base
        
        elif any(k in self.catNbr for k in ("699", "799")):
            load = min(base * eff_enroll, 5.0)

        else:
            rate = self._adjustForEnrollment(base)
            load = self.maxUnits * rate
            if "lecture" in self.courseCategory:
                load = min(load, self.maxUnits * (20.0/3.0))

        if any(code in self.classCat for code in self.special):
            load += self.maxUnits * self.policy.get("specialCoursesRate", 0.005)

        self.load = load
        return round(load, 2)

    def adjustLoadDivision(self, d):
        if d <= 1:
            return self.load
        if self.load is None:
            self.load = self.calculateLoad()
        self.load /= d 
        return self.load

# ---------------------------------------------------------------------------
# Faculty container
# ---------------------------------------------------------------------------

class FacultyMember:
    def __init__(self, name, email, emplid, initialRole, track=None):
        self.name = name 
        self.email = email 
        self.emplid = int(emplid)
        self.roles = {initialRole} 
        self.courses = {} 
        self.totalLoad = 0.0 
        self.track = track

    def addCourse(self, course):
        self.courses[course.getGroupKeyForGrouping()] = course 
        self.roles.add(course.instructorRole)

    def calculateTotalLoad(self):
        self.totalLoad = sum(c.calculateLoad() for c in self.courses.values()) 
        return self.totalLoad

# ---------------------------------------------------------------------------
# Co‑convened adjustment
# ---------------------------------------------------------------------------

def adjust_co_convened(courses: Iterable[Course]) -> None:
    bundles: Dict[Tuple, List[Course]] = defaultdict(list)
    for c in courses:
        if c.instructorEmplid is None:
            continue
        
        if not all(c._meeting_signature()):
            continue

        key = c.getGroupKeyForCollapsing()
        bundles[key].append(c)

    for same in bundles.values():
        if len(same) <= 1: 
            continue
        
        ids = [
            f"{c.rawData.get('Subject','').strip()} {c.catNbr}-{c.rawData.get('Section','').strip()}"
            for c in same
        ]

        for c in same:
            me = f"{c.rawData.get('Subject','').strip()} {c.catNbr}-{c.rawData.get('Section','').strip()}"
            c.co_convened_members = [other for other in ids if other != me]
        
        rep, *others = sorted(same, key=lambda c: c.maxUnits, reverse=True)
        combined = sum(c.enrollTotal for c in same)
        rep.enrollTotal = combined
        rep.load = None
        rep.load = rep.calculateLoad()

        for extra in others:
            extra.load = 0


################################################################################
'''
def main():
            # 1) Load raw data
            converters = {
                'Max Units': lambda x: float(x) if pd.notna(x) else 0.0,
                'Enroll Total': lambda x: int(float(x)) if pd.notna(x) else 0
            }
            raw_df = pd.read_excel("FIle 1 choke a goat.xlsx", sheet_name='Raw Data', converters=converters)
            raw_df = raw_df[raw_df.apply(rowIsValid, axis=1)].reset_index(drop=True)
            raw_df = raw_df.drop_duplicates(subset=[
                'Instructor Emplid', 'Term', 'Subject', 'Cat Nbr', 'Section',
                'Start Date', 'End Date', 'Start Time', 'End Time', 'Facility Building', 'Facility Room', 'Days'
            ])
            
            # 2) Supporting data
            policy = loadWorkloadPolicy("workload_policy.xlsx")
            tracks = loadInstructorTrack("FIle 1b CC or TT.xlsx")
            special = loadSpecialCourses("CEFNS courses with extra load assigned.xlsx")

            # 3) Build structures
            faculty = {}
            courseGroups = {}
            other = {}

            for _, row in raw_df.iterrows():
                role = str(row.get('Instructor Role', '')).strip().upper()

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
                valid = [c for c in lst if all(c._meeting_signature())]
                pi_only = [c for c in valid if c.instructorRole.upper()=="PI"]
                unique_emplids = {c.instructorEmplid for c in pi_only}
                if len(unique_emplids) >= 2:
                    names = [c.rawData.get('Instructor','').strip() for c in pi_only]
                    for c in pi_only:
                        c.team_taught_members = [
                            n for n in names 
                            if n != c.rawData.get('Instructor','').strip()
                        ]
                        c.adjustLoadDivision(len(unique_emplids))

            # 5) Co‑convened adjustment
            adjust_co_convened([c for lst in courseGroups.values() for c in lst])

            # 6) Calculate summary
            summary_rows = []
            for fac in faculty.values():
                fac.calculateTotalLoad()
                units = sorted({getattr(c, 'unit', '') for c in fac.courses.values() if getattr(c, 'unit', '')})

                course_list = []
                for c in fac.courses.values():
                    # base label
                    subject = c.rawData.get('Subject','').strip()
                    section = c.rawData.get('Section','').strip()
                    desc    = c.rawData.get('Class Description','').strip().title()
                    label   = f"{subject} {c.catNbr}-{section} – {desc}"

                    # tag on any partner info
                    if getattr(c, 'co_convened_members', None):
                        label += f" (co‑convened with {', '.join(c.co_convened_members)})"
                    if getattr(c, 'team_taught_members', None):
                        label += f" (team‑taught with {', '.join(c.team_taught_members)})"

                    course_list.append(label)

                course_list.sort()
                summary_rows.append({
                    'Instructor': fac.name,
                    'Emplid': fac.emplid,
                    'Track': fac.track or 'Unknown',
                    'Total Workload': round(fac.totalLoad, 2),
                    'Units Taught': ', '.join(units),
                    'Courses Taught': '; '.join(course_list)
                })

if __name__ == "__main__":
    main()

'''