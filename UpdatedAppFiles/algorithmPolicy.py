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
        data.get("Start Time"),
        _norm(data.get("Facility Building")),
        _norm(data.get("Facility Room")),
    )

# ---------------------------------------------------------------------------
# Policy / loaders
# ---------------------------------------------------------------------------

def loadWorkloadPolicy(path: str | None = None) -> dict:
    defaults = {
        "independentStudyRate": 0.33,
        "laboratoryRate": 4.17,
        "lectureRate": 3.33,
        "thesisRate": 3.33,
        "maxLoadCap": 5.0,
        "thesisCap": 13.33,
        "lectureThreshold": {"low": 90, "mid": 150, "high": 200},
        "midRate": 4.17,
        "highRate": 5.0,
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
    required = ["Course Category (CCAT)", "Max Units", "Enroll Total", "Instructor Role", "Instructor Emplid"]
    if any(pd.isna(row.get(c)) for c in required):
        return False
    ccat = _norm(row.get("Course Category (CCAT)"))
    exempt = {"independent study", "research", "thesis", "dissertation", "fieldwork"}
    if any(k in ccat for k in exempt):
        return True
    meeting_cols = ["Start Date", "Start Time", "Facility Building", "Facility Room"]
    return not any(pd.isna(row.get(c)) for c in meeting_cols)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

GRADING_INTENSIVE_KEYWORDS = {"capstone", "writing intensive", "w-intensive"}
FIELD_TRIP_KEYWORDS = {"field trip"}
FIELD_TRIP_BONUS_WLU = 0.15
SPECIAL_COURSES_EXTRA_RATE = 0.005

# ---------------------------------------------------------------------------
# Course object
# ---------------------------------------------------------------------------

class Course:
    def __init__(self, data: dict, policy: dict, special: Set[str]):
        self.rawData = data
        self.policy = policy
        self.special = special

        self.courseCategory = _norm(data.get("Course Category (CCAT)"))
        self.classDescription = _norm(data.get("Class Description"))
        cat_raw = str(data.get("Cat Nbr", "")).strip()
        self.catNbr = str(int(float(cat_raw))) if cat_raw.replace(".", "", 1).isdigit() else cat_raw
        self.instructorRole = _norm(data.get("Instructor Role"))

        self.maxUnits = float(data.get("Max Units", 0) or 0)
        self.enrollTotal = int(data.get("Enroll Total", 0) or 0)
        emplid_raw = data.get("Instructor Emplid", None)
        self.instructorEmplid = int(float(emplid_raw)) if pd.notna(emplid_raw) else None

        self.startDate = data.get("Start Date")
        self.startTime = data.get("Start Time")
        self.facilityBuilding = data.get("Facility Building")
        self.facilityRoom = data.get("Facility Room")

        self.isGradingIntensive = any(k in self.classDescription for k in GRADING_INTENSIVE_KEYWORDS) or self.catNbr.endswith("W")
        self.isThesis = (
            any(k in self.courseCategory for k in ("thesis", "dissertation")) or
            self.catNbr in {"699", "799"}
        )
        self.hasFieldTrip = any(k in self.classDescription for k in FIELD_TRIP_KEYWORDS)

        self.unit = str(data.get("Unit", "")).strip()
        self.load: float | None = None

    # ------------------------------------------------------------------
    def _meeting_signature(self):
        return _meeting_signature(self.rawData)

    def getGroupKeyForGrouping(self):
        term = self.rawData.get("Term")
        subject = self.rawData.get("Subject")
        section = self.rawData.get("Section")
        return (term, subject, self.catNbr, section) + self._meeting_signature()

    def getGroupKeyForCollapsing(self):
        term = self.rawData.get("Term")
        subject = self.rawData.get("Subject")
        section = self.rawData.get("Section")
        classNbr = self.rawData.get("Class Nbr")
        return (self.instructorEmplid, term, subject, self.catNbr, section, classNbr) + self._meeting_signature()

    # ------------------------------------------------------------------
    def _baseRate(self):
        p = self.policy
        if self.isThesis:
            return float(p.get("thesisRate", 3.33))
        if any(k in self.courseCategory for k in ("independent study", "research", "fieldwork")):
            return float(p.get("independentStudyRate", 0.33))
        if "laboratory" in self.courseCategory:
            return float(p.get("laboratoryRate", 4.17))
        return float(p.get("lectureRate", 3.33))

    def _adjustForEnrollment(self, base):
        if "lecture" not in self.courseCategory or self.isGradingIntensive:
            return base
        p = self.policy; low, mid, high = (float(p["lectureThreshold"][k]) for k in ("low", "mid", "high"))
        s = self.enrollTotal
        if s < low:
            return base
        if low <= s <= mid:
            return float(p["lectureRate"] + (s - low)/(mid - low)*(p["midRate"] - p["lectureRate"]))
        if mid < s <= high:
            return float(p["midRate"] + (s - mid)/(high - mid)*(p["highRate"] - p["midRate"]))
        return float(p["highRate"])

    # ------------------------------------------------------------------
    def calculateLoad(self):
        if self.enrollTotal == 0:
            return 0.0

        base = self._baseRate()
        eff_enroll = min(self.enrollTotal, 24) if self.isGradingIntensive else self.enrollTotal

        if self.isThesis or any(k in self.courseCategory for k in ("independent study", "research", "fieldwork")):
            load = self.maxUnits * (base * eff_enroll)
            cap = self.policy.get("thesisCap", 13.33) if self.isThesis else self.policy.get("maxLoadCap", 5.0)
            load = min(load, cap)
        else:
            rate = self._adjustForEnrollment(base)
            load = self.maxUnits * rate
            if "lecture" in self.courseCategory:
                load = min(load, self.maxUnits * (20.0/3.0))

        if any(code in self.classDescription for code in self.special):
            load += self.maxUnits * SPECIAL_COURSES_EXTRA_RATE
        if self.hasFieldTrip:
            load += FIELD_TRIP_BONUS_WLU

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
        self.name = name; self.email = email; self.emplid = int(emplid)
        self.roles = {initialRole}; self.courses = {}; self.totalLoad = 0.0; self.track = track

    def addCourse(self, course):
        self.courses[course.getGroupKeyForGrouping()] = course; self.roles.add(course.instructorRole)

    def calculateTotalLoad(self):
        self.totalLoad = sum(c.calculateLoad() for c in self.courses.values()); return self.totalLoad

# ---------------------------------------------------------------------------
# Co‑convened adjustment
# ---------------------------------------------------------------------------

def adjust_co_convened(courses: Iterable[Course], mode: str = "collapse") -> None:
    bundles: Dict[Tuple, List[Course]] = defaultdict(list)
    for c in courses:
        if c.instructorEmplid is None:
            continue
        key = c.getGroupKeyForCollapsing()
        bundles[key].append(c)

    for same in bundles.values():
        if len(same) <= 1: continue
        if mode == "collapse":
            combined = sum(c.enrollTotal for c in same)
            rep = same[0]; rep.enrollTotal = combined; rep.calculateLoad()
            for extra in same[1:]: extra.load = 0.00
        elif mode == "split":
            for c in same: c.calculateLoad()
            share = 1/len(same)
            for c in same: c.load *= share
        else:
            raise ValueError("mode must be 'collapse' or 'split'")