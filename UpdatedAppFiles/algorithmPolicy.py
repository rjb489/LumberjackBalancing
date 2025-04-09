
import pandas as pd
from collections import defaultdict
from typing import Iterable, Dict, List, Tuple, Set

###############################################################################
# Updated workload algorithm (2025‑04‑08)
###############################################################################
# Handles all student‑related activities that can be inferred from the Peoplesoft
# schedule export while restricting rows to Primary Instructors (PI).
###############################################################################

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------
def _norm(s):
    """Normalise strings for case‑insensitive comparison and handle NaN."""
    return str(s).strip().lower() if pd.notna(s) else ""

def _meeting_signature(row_or_course) -> Tuple:
    """Return (date, time, building, room) tuple identifying a meeting."""
    data = row_or_course.rawData if hasattr(row_or_course, "rawData") else row_or_course
    return (
        data.get("Start Date"),
        data.get("Start Time"),
        _norm(data.get("Facility Building")),
        _norm(data.get("Facility Room")),
    )

# ---------------------------------------------------------------------------
# Policy / supporting‑data loaders
# ---------------------------------------------------------------------------
def loadWorkloadPolicy(path: str | None = None) -> dict:
    """Load workload policy spreadsheet or fall back to defaults."""
    defaults = {
        "independentStudyRate": 0.33,
        "laboratoryRate": 4.17,
        "lectureRate": 3.33,
        "thesisRate": 3.33,
        "maxLoadCap": 5.0,
        "lectureThreshold": {"low": 90, "mid": 150, "high": 200},
        "midRate": 4.17,
        "highRate": 5.0,
    }
    policy = defaults.copy()
    if path:
        try:
            df = pd.read_excel(path)
            kv = {str(k).strip(): v for k, v in zip(df.iloc[:, 0], df.iloc[:, 1])}
            # flatten lectureThreshold_* keys
            if all(k in kv for k in ("lectureThreshold_low", "lectureThreshold_mid", "lectureThreshold_high")):
                policy["lectureThreshold"] = {
                    "low": float(kv.pop("lectureThreshold_low")),
                    "mid": float(kv.pop("lectureThreshold_mid")),
                    "high": float(kv.pop("lectureThreshold_high")),
                }
            # merge rest
            for k, v in kv.items():
                try:
                    policy[k] = float(v)
                except (TypeError, ValueError):
                    policy[k] = v
        except Exception as e:
            print("Warning: failed to load policy file – using defaults:", e)
    return policy

def loadInstructorTrack(path: str) -> dict:
    try:
        df = pd.read_excel(path)
        return pd.Series(df["Track"].values, index=df["Instructor Emplid"]).to_dict()
    except Exception as e:
        print("Warning: failed to load instructor track file:", e)
        return {}

def loadSpecialCourses(path: str) -> set:
    try:
        df = pd.read_excel(path)
        return set(df["Course"].dropna().astype(str).str.strip().str.lower())
    except Exception as e:
        print("Warning: failed to load special courses file:", e)
        return set()

# ---------------------------------------------------------------------------
# Row filter
# ---------------------------------------------------------------------------
def rowIsValid(row: pd.Series) -> bool:
    """Return True if row contains required info for workload calc."""
    required = ["Course Category (CCAT)", "Max Units", "Enroll Total",
                "Instructor Role", "Instructor Emplid"]
    if any(pd.isna(row.get(c)) for c in required):
        return False
    ccat = _norm(row.get("Course Category (CCAT)"))

    exempt = {"independent study", "research", "thesis", "dissertation", "fieldwork"}
    if any(k in ccat for k in exempt):
        return True  # exempt from meeting info requirement

    meeting_cols = ["Start Date", "Start Time", "Facility Building", "Facility Room"]
    return not any(pd.isna(row.get(c)) for c in meeting_cols)

# ---------------------------------------------------------------------------
# Constants for special detection
# ---------------------------------------------------------------------------
GRADING_INTENSIVE_KEYWORDS = {"capstone", "writing intensive", "w-intensive"}
FIELD_TRIP_KEYWORDS = {"field trip"}
FIELD_TRIP_BONUS_WLU = 0.15                # flat WLU bump
SPECIAL_COURSES_EXTRA_RATE = 0.005         # extra per‑unit bump
COMMITTEE_MEMBER_WLU = 0.30     # 1 % of AY = 0.30 WLU

# ---------------------------------------------------------------------------
# Course object
# ---------------------------------------------------------------------------
class Course:
    def __init__(self, data: dict, policy: dict, specialCourses: Set[str]):
        self.rawData = data
        self.policy = policy
        self.specialCourses = specialCourses

        # Basic attributes
        self.courseCategory = _norm(data.get("Course Category (CCAT)")) or ""
        self.classDescription = _norm(data.get("Class Description")) or ""
        self.catNbr = str(data.get("Cat Nbr", "")).strip()
        self.instructorRole = _norm(data.get("Instructor Role")) or ""

        # Numeric attributes
        self.maxUnits = float(data.get("Max Units", 0) or 0)
        self.enrollTotal = int(data.get("Enroll Total", 0) or 0)
        self.instructorEmplid = (
            int(float(data.get("Instructor Emplid", 0)))
            if pd.notna(data.get("Instructor Emplid"))
            else None
        )

        # Meeting info
        self.startDate = data.get("Start Date")
        self.startTime = data.get("Start Time")
        self.facilityBuilding = data.get("Facility Building")
        self.facilityRoom = data.get("Facility Room")

        # Flags
        self.isGradingIntensive = any(k in self.classDescription for k in GRADING_INTENSIVE_KEYWORDS) or self.catNbr.endswith("W")
        self.isThesis = any(k in self.courseCategory for k in ("thesis", "dissertation")) or self.catNbr in {"699", "799"}
        self.hasFieldTrip = any(k in self.classDescription for k in FIELD_TRIP_KEYWORDS)

        self.unit = str(data.get("Unit", "")).strip()
        self.load: float | None = None

    # ---------------- grouping helpers ----------------
    def _meeting_signature(self):
        return _meeting_signature(self.rawData)

    def getGroupKey(self) -> Tuple:
        """Key for team‑taught grouping (same course & meeting)."""
        if any(k in self.courseCategory for k in ("research", "thesis", "dissertation")):
            return (self.instructorEmplid,
                    self.startDate,
                    self.rawData.get("Term"),
                    self.rawData.get("Subject"),
                    "research")
        term = _norm(self.rawData.get("Term"))
        subject = _norm(self.rawData.get("Subject"))
        section = _norm(self.rawData.get("Section"))
        return (term, subject, self.catNbr.lower(), section) + self._meeting_signature()

    # ---------------- rate helpers ----------------
    def _baseRate(self) -> float:
        p = self.policy
        if self.isThesis:
            return float(p.get("thesisRate", 3.33))
        if any(k in self.courseCategory for k in ("independent study", "research", "fieldwork")):
            return float(p.get("independentStudyRate", 0.33))
        if "laboratory" in self.courseCategory:
            return float(p.get("laboratoryRate", 4.17))
        return float(p.get("lectureRate", 3.33))

    def _adjustForEnrollment(self, base: float) -> float:
        if "lecture" not in self.courseCategory or self.isGradingIntensive:
            return base
        p = self.policy
        low, mid, high = (float(p["lectureThreshold"][k]) for k in ("low", "mid", "high"))
        s = self.enrollTotal
        if s < low:
            return base
        if low <= s <= mid:
            return float(p["lectureRate"] + (s - low)/(mid - low)*(p["midRate"] - p["lectureRate"]))
        if mid < s <= high:
            return float(p["midRate"] + (s - mid)/(high - mid)*(p["highRate"] - p["midRate"]))
        return float(p["highRate"])

    # ---------------- public API ----------------
    def calculateLoad(self) -> float:
        if self.load is not None:
            return self.load

        if self.enrollTotal == 0:
            self.load = 0.0
            return self.load

        base = self._baseRate()
        effective_enroll = min(self.enrollTotal, 24) if self.isGradingIntensive else self.enrollTotal

        if self.isThesis or any(k in self.courseCategory for k in ("independent study", "research", "fieldwork")):
            load = self.maxUnits * (base * effective_enroll)
            load = min(load, self.policy.get("maxLoadCap", 5.0))
        else:
            rate = self._adjustForEnrollment(base)
            load = self.maxUnits * rate
            if "lecture" in self.courseCategory:
                load = min(load, self.maxUnits * (20.0/3.0))  # 6.67 WLU/credit cap

        # Special course bumps
        if any(code in self.classDescription for code in self.specialCourses):
            load += self.maxUnits * SPECIAL_COURSES_EXTRA_RATE
        if self.hasFieldTrip:
            load += FIELD_TRIP_BONUS_WLU

        self.load = load
        return load

    def adjustLoadDivision(self, divisor: int) -> float:
        if self.load is None:
            self.calculateLoad()
        self.load /= divisor
        return self.load

    def __repr__(self):
        return f"<Course {self.getGroupKey()}, load: {self.load:.2f}>"

# ---------------------------------------------------------------------------
# Faculty container
# ---------------------------------------------------------------------------
class FacultyMember:
    def __init__(self, name: str, email: str, emplid: int, initialRole: str, track: str | None = None):
        self.name = name
        self.email = email
        self.emplid = int(emplid)
        self.roles = {initialRole}
        self.courses: Dict[Tuple, Course] = {}
        self.totalLoad = 0.0
        self.track = track

    def addCourse(self, course: Course) -> None:
        self.courses[course.getGroupKey()] = course
        self.roles.add(course.instructorRole)

    def calculateTotalLoad(self) -> float:
        self.totalLoad = sum(c.calculateLoad() for c in self.courses.values())
        return self.totalLoad

# ---------------------------------------------------------------------------
# Co‑convened adjustment
# ---------------------------------------------------------------------------
def adjust_co_convened(courses: Iterable[Course], mode: str = "collapse") -> None:
    """Collapse or split loads for cross‑listed courses taught simultaneously."""
    bundles: Dict[Tuple, List[Course]] = defaultdict(list)
    for c in courses:
        if c.instructorEmplid is None:
            continue
        bundles[(c.instructorEmplid, c._meeting_signature())].append(c)

    for same_meeting in bundles.values():
        if len(same_meeting) <= 1:
            continue

        if mode == "collapse":
            combined_enroll = sum(c.enrollTotal for c in same_meeting)
            rep = same_meeting[0]
            rep.enrollTotal = combined_enroll
            rep.calculateLoad()
            for extra in same_meeting[1:]:
                extra.load = COMMITTEE_MEMBER_WLU
        elif mode == "split":
            for c in same_meeting:
                c.calculateLoad()
            share = 1.0 / len(same_meeting)
            for c in same_meeting:
                c.load *= share
        else:
            raise ValueError("mode must be 'collapse' or 'split'")
