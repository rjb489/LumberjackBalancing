import pandas as pd
import numpy as np

def loadWorkloadPolicy(filePath: str = None) -> dict:
   defaultPolicy = {
      'independentStudyRate': 0.33,
      'laboratoryRate': 4.17,
      'lectureRate': 3.33,
      'maxLoadCap': 5.0,
      'lectureThreshold': {
         'low': 90,
         'mid': 150,
         'high': 200,
      },
      'midRate': 4.17,
      'highRate': 5.0
   }

   policy = defaultPolicy.copy()

   if filePath:
      try:
         df = pd.read_excel(filePath)
         policyDict = {k: float(v) if isinstance(v, (int, float, str)) and v not in [None, ''] else v 
                          for k, v in dict(zip(df.iloc[:, 0], df.iloc[:, 1])).items()}
         
         if('lectureThreshold_low' in policyDict and
            'lectureThreshold_mid' in policyDict and
            'lectureThreshold_high' in policyDict):
            policyDict['lectureThreshold'] = {
               'low': float(policyDict.pop('lectureThreshold_low')),
               'mid': float(policyDict.pop('lectureThreshold_mid')),
               'high': float(policyDict.pop('lectureThreshold_high'))
            }

         for key, value in policyDict.items():
            if isinstance(value, dict):
               policy[key] = value
            
            else:
               try:
                  policy[key] = float(value)
               
               except (ValueError, TypeError):
                  policy[key] = value
         
         return policy
      
      except Exception as e:
         print("Error loading policy file, using default policy values:", e)
   
   return policy

def loadInstructorTrack(filePath: str) -> dict:
   try:
      df = pd.read_excel(filePath)
      trackDict = pd.Series(df['Track'].values, index=df['Instructor Emplid']).to_dict()
      return trackDict
   except Exception as e:
      print("Error loading instructor track file:", e)
      return {}

def loadSpecialCourses(filePath: str) -> set:
   try:
      df = pd.read_excel(filePath)
      specialSet = set(df["Course"].dropna().astype(str).str.strip().str.lower())
      return specialSet
   except Exception as e:
      print("Error loading special courses file:", e)
      return set()

################################################################################

def rowIsValid(row: pd.Series) -> bool:
   common = ['Course Category (CCAT)', 'Max Units', 'Enroll Total', 'Instructor Role', 'Instructor Emplid']
   for col in common:
      if pd.isna(row.get(col)):
         return False
   
   courseCategory = str(row.get('Course Category (CCAT)', '')).lower()
   exemptKeywords = ["independent study", "research", "thesis", "dissertation", "fieldwork"]
   
   if any(keyword in courseCategory for keyword in exemptKeywords):
        return True

   else:
      meeting = ['Start Date', 'Start Time', 'Facility Room', 'Facility Building']
      for col in meeting:
         if pd.isna(row.get(col)):
            return False
         
      return True

################################################################################
SPECIAL_COURSES_EXTRA_RATE = 0.005

class Course:
   def __init__(self, data: dict, policyParams: dict, specialCourses: set):
      self.rawData = data

      self.courseCategory = str(data.get('Course Category (CCAT)', '')).strip().lower()

      try:
         self.maxUnits = float(data.get('Max Units', 0))
      except(ValueError, TypeError):
         self.maxUnits = 0.0
      
      try:
         self.enrollTotal = int(data.get('Enroll Total', 0))
      except(ValueError, TypeError):
         self.enrollTotal = 0
      
      self.classDescription = str(data.get('Class Description', '')).strip().lower()
      self.instructorRole = str(data.get('Instructor Role', '')).strip().upper()
      self.startDate = data.get('Start Date', None)
      self.startTime = data.get('Start Time', None)
      self.facilityRoom = str(data.get('Facility Room', '')).strip()
      self.facilityBuilding = str(data.get('Facility Building', '')).strip()
      self.unit = str(data.get('Unit', '')).strip()
      self.policyParams = policyParams
      self.specialCourses = specialCourses
      self.load = None

   def getGroupKey(self) -> tuple:
      term = self.rawData.get('Term')
      subject = self.rawData.get('Subject')
      catNbr = self.rawData.get('Cat Nbr')
      section = self.rawData.get('Section')
      instructor = self.rawData.get('Instructor Emplid')

      if any(keyword in self.courseCategory for keyword in ["research", "thesis", "dissertation"]):
         return (instructor, self.startDate, term, subject, "research")
      
      else:
         if term and subject and catNbr and section:
            return (str(term).strip().lower(),
                    str(subject).strip().lower(),
                    str(catNbr).strip().lower(),
                    str(section).strip().lower())
         else:
            return (self.startDate, self.instructorRole)
          
   def getBaseRate(self) -> float:
      policyRate = self.policyParams

      if any(keyword in self.courseCategory for keyword in ["independent study", 
                            "research", "thesis", "dissertation", "fieldwork"]):
         return float(policyRate.get('independentStudyRate', 0.33))
      
      elif "laboratory" in self.courseCategory:
         return float(policyRate.get('laboratoryRate', 4.17))
      
      elif any(keyword in self.courseCategory for keyword in ["recitation", 
                                                         "seminar", "lecture"]):
         return float(policyRate.get('lectureRate', 3.33))
      
      else:
         return float(policyRate.get('lectureRate', 3.33))
      
   def adjustForEnrollment(self, baseRate: float) -> float:
      studentNumber = self.enrollTotal
      policyRate = self.policyParams
      thresholds = policyRate.get('lectureThreshold', {'low': 90, 'mid': 150, 'high': 200})

      low = float(thresholds.get('low', 90))
      mid = float(thresholds.get('mid', 150))
      high = float(thresholds.get('high', 200))

      if "lecture" in self.courseCategory:
         if studentNumber < low:
            return baseRate
         
         elif low <= studentNumber <= mid:
            rate = float(policyRate.get('lectureRate', 3.33) + ((studentNumber - thresholds.get('low', 90)) /
                   (thresholds.get('mid', 150) - thresholds.get('low', 90))) * (policyRate.get('midRate', 4.17) - policyRate.get('lectureRate', 3.33)))
            return rate
         
         elif mid < studentNumber <= high:
            rate = float(policyRate.get('midRate', 4.17) + ((studentNumber - thresholds.get('mid', 150)) /
                      (thresholds.get('high', 200) - thresholds.get('mid', 150))) * (policyRate.get('highRate', 5.0) - policyRate.get('midRate', 4.17)))
            return rate
         
         else:
            return float(policyRate.get('highRate', 5.0))
         
      else:
         return baseRate
      
   def calculateLoad(self) -> float:
      if self.enrollTotal == 0:
         self.load = 0.0
         return self.load

      policyRate = self.policyParams

      if any(keyword in self.courseCategory for keyword in ["independent study", 
                            "research", "thesis", "dissertation", "fieldwork"]):
         baseRate = self.getBaseRate()
         load = self.maxUnits * (baseRate * self.enrollTotal)
         load = min(load, policyRate.get('maxLoadCap', 5.0))

      else:
         baseRate = self.getBaseRate()
         adjustedRate = self.adjustForEnrollment(baseRate)
         load = self.maxUnits * adjustedRate 

         if "lecture" in self.courseCategory:
            maxRate = (20.0 / 3.0)
            maxload = self.maxUnits * maxRate
            load = min(load, maxload)
      
      extraLoad = 0.0

      for specialCode in self.specialCourses:
         if specialCode in self.classDescription:
            extraLoad = self.maxUnits * SPECIAL_COURSES_EXTRA_RATE
            break
      
      self.load = load + extraLoad
      return self.load
   
   def adjustLoadDivision(self, count: int) -> float:
      if self.load is None:
         self.calculateLoad()
      self.load = self.load / count
      return self.load
   
   def __repr__(self):
        return f"<Course {self.getGroupKey()}, load: {self.load:.2f}>"

################################################################################

class FacultyMember:
   def __init__(self, name: str, email: str, emplid: int, initialRole: str, track: str = None):
      self.name = name
      self.email = email
      self.emplid = int(emplid)
      self.roles = {initialRole}
      self.courses = {}
      self.totalLoad = 0.0
      self.track = track
      self.trackPercentage = None

   def addCourse(self, course: Course) -> None:
      key = course.getGroupKey()
      if key not in self.courses:
         self.courses[key] = course
      self.roles.add(course.instructorRole)
   
   def calculateTotalLoad(self) -> float:
      total = 0.0
      for course in self.courses.values():
         if course.load is not None:
            total += course.load

         else:   
            total += course.calculateLoad()

      self.totalLoad = total
      return total

   def calculatePercentage(self) -> float:
      if self.track and self.track.upper() == "CT":
         self.trackPercentage = (self.totalLoad / 40.0) * 100
      
      else:
         self.trackPercentage = (self.totalLoad / 30.0) * 100
      
      return self.trackPercentage

################################################################################
# MAIN PROCESSING PORTION OF ALGORITHM
################################################################################

def main():
   converters = {
    'Max Units': lambda x: float(x) if pd.notna(x) else 0.0,
    'Enroll Total': lambda x: int(float(x)) if pd.notna(x) else 0
   }

   rawCourseDataFile = "FIle 1 choke a goat.xlsx"
   fileSheetName = "Raw Data"
   rawCourseData = pd.read_excel(rawCourseDataFile, sheet_name=fileSheetName, converters=converters)
   rawCourseData = rawCourseData[rawCourseData.apply(rowIsValid, axis=1)].reset_index(drop=True)

   policyFile = "workload_policy.xlsx"
   policyParams = loadWorkloadPolicy(policyFile)

   instructorTrackFile = "FIle 1b CC or TT.xlsx"
   instructorTracks = loadInstructorTrack(instructorTrackFile)

   specialCoursesFile= "CEFNS courses with extra load assigned.xlsx"
   specialCourses = loadSpecialCourses(specialCoursesFile)

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
      print(f"Faculty: {faculty.name}, (ID: {faculty.emplid})")
      if faculty.track and faculty.track.upper() == "CT":
         print(f"{faculty.totalLoad:.2f}\n   CT Percentage (baseline 40%): {percentage:.2f}%\n")
      else:
         print(f"{faculty.totalLoad:.2f}\n   TT Percentage (baseline 30%): {percentage:.2f}%\n")
      print(f"Courses: {faculty.courses}\n")

if __name__ == "__main__":
   main()
