import pandas as pd
import numpy as np

# Load data from file (Hard coded, change to user input in final version)
courseData = pd.read_excel("Scot_Randomized.xlsx")

# Definition of essential categories necessary for data calculation
# Future version should include building for composite key
essentialColumns = [
    'Course Category (CCAT)', 
    'Max Units', 
    'Enroll Total', 
    'Instructor Role', 
    'Instructor Emplid', 
    'Start Date', 
    'Start Time', 
    'Facility Room'
]

# Clean data by removing rows with missing values, duplicates, and reset index
courseData = courseData.dropna(subset=essentialColumns)
courseData = courseData.drop_duplicates()
courseData = courseData.reset_index(drop=True)

# Definition of course object
class Course:
   def __init__(self, data):
      self.courseCategory = str(data.get('Course Category (CCAT)', '')).strip().lower()

      try:
         self.maxUnits = float(data.get('Max Units', 0))
      except(ValueError, TypeError):
         self.maxUnits = 0.0
      
      try:
         self.enrollTotal = int(data.get('Enroll Total', 0))
      except(ValueError, TypeError):
         self.enrollTotal = 0
      
      self.classDescription = data.get('Class Description', '')
      self.instructorRole = str(data.get('Instructor Role', '')).strip().upper()
      self.startDate = data.get('Start Date', None)
      self.startTime = data.get('Start Time', None)
      self.facilityRoom = str(data.get('Facility Room', '')).strip()
      self.load = None

   # Future version should include building
   def getCompositeKey(self):
      return (self.startDate, self.startTime, self.facilityRoom)
   
   def getBaseRate(self):
      if any(keyword in self.courseCategory for keyword in ["independent study", 
                            "research", "thesis", "dissertation", "fieldwork"]):
         return 0.33
      
      elif "laboratory" in self.courseCategory:
         return 4.17
      
      elif any(keyword in self.courseCategory for keyword in ["recitation", 
                                                         "seminar", "lecture"]):
         return 3.33
      
      else:
         return 3.33
      

   def adjustForEnrollment(self, baseRate):
      studentNumber = self.enrollTotal

      if "lecture" in self.courseCategory:
         if studentNumber < 90:
            return baseRate
         elif 90 <= studentNumber <= 150:
            rate = 3.33 + ((studentNumber - 90) / 60.0) * (4.17 - 3.33)
            return rate
         elif 150 < studentNumber <= 200:
            rate = 4.17 + ((studentNumber - 150) / 50.0) * (5.0 - 4.17)
            return rate
         else:
            return 5.0
      else:
         return baseRate
      
   def calculateLoad(self):
      if any(keyword in self.courseCategory for keyword in ["independent study", 
                            "research", "thesis", "dissertation", "fieldwork"]):
         baseRate = self.getBaseRate()
         load = self.maxUnits * (baseRate * self.enrollTotal)
         load = min(load, 5.0)
      else:
         baseRate = self.getBaseRate()
         adjustedRate = self.adjustForEnrollment(baseRate)
         load = self.maxUnits * adjustedRate 

         if "lecture" in self.courseCategory:
            maxRate = (20.0 / 3.0)
            maxload = self.maxUnits * maxRate
            load = min(load, maxload)
      
      self.load = load
      return load

# Define faculty member object
class FacultyMember:
   def __init__(self, name, email, emplid, initialRole):
      self.name = name
      self.email = email
  
      self.emplid = int(float(emplid))
      self.roles = {initialRole}
      self.courses = {}
      self.totalLoad = 0.0

   def addCourse(self, course):
      key = course.getCompositeKey()
      if key not in self.courses:
         self.courses[key] = course
      self.roles.add(course.instructorRole)
   
   def calculateTotalLoad(self):
      total = 0.0
      for course in self.courses.values():
         if course.load is not None:
            total += course.load
         else:   
            total += course.calculateLoad()
      self.totalLoad = total
      return total

   def calculatePercentage(self):
      # IMPORTANT: Must be used after calculateTotalLoad
      # ttvalue expected percentage = 30% per semester = 60% AY
      # ctvalue expected percentage = 40% per semester = 80% AY
      ttValue = (self.totalLoad / 30.0) * 100
      ctValue = (self.totalLoad / 40.0) * 100
      return ttValue, ctValue

################################################################################
# MAIN PROCESSING PORTION OF ALGORITHM
################################################################################

# Define dictionaries to store faculty and other staff
facultyDict = {}
otherStaff = {}

# Define composite map to store courses with same composite key 
# Used for Co-convened and Team-taught courses
compositeMap = {}

# Iterate through course data, extract faculty information
for index, row in courseData.iterrows():

   # Extract faculty role
   role = str(row.get('Instructor Role', '')).strip().upper()
   
   # IMPORTANT: current version only uses PI (primary instructor) role
   # other roles are stored separately
   # some cases are not fully accounted for (labs only calculated for PI role for example)

   # Check if faculty is PI, if not store in other staff dictionary and skip to next row
   if role != "PI":
      emplid = str(row.get('Instructor Emplid', '')).strip()
      if emplid not in otherStaff:
         otherStaff[emplid] = []
      otherStaff[emplid].append(row.to_dict())
      continue
   
   # if PI, extract faculty information
   emplid = int(float(row.get('Instructor Emplid', 0)))
   name = row.get('Instructor', '')
   email = row.get('Instructor Email', '')

   # Create course object
   courseObject = Course(row.to_dict())
   # Generate composite key
   compKey = courseObject.getCompositeKey()
   # If composite key not in map already, add to map
   # necessary for Co-convened and Team-taught courses 
   if compKey not in compositeMap:
      compositeMap[compKey] = []
   # Append faculty and course to composite map (if same composite key, multiple items added to same key)
   compositeMap[compKey].append((emplid, courseObject))

   # If faculty not in dictionary, generate faculty object
   if emplid not in facultyDict:
      facultyDict[emplid] = FacultyMember(name, email, emplid, role)
   # Add course to current faculty member
   facultyDict[emplid].addCourse(courseObject)

# Adjust the load for Co-convened and Team-taught courses
for compKey, courseList in compositeMap.items():
   # check if current key has more than one course
   # this means that either two faculty members have the same course assigned to them (team-taught)
   # OR the course composite key is the same for two different courses (co-convened)
   if len(courseList) > 1:
      # use number of courses in key to adjust load
      # indicates the number of faculty members teaching 
      # OR number of different courses with same composite key (should be 2 for undergrad and grad co-convened courses)
      count = len(courseList)

      # iterate through courses in current key
      for emplid, courseObject in courseList:
         # calculate original load of course then divide by count
         originalLoad = courseObject.calculateLoad()
         adjustedLoad = originalLoad / count

         # set adjusted load for current course object
         # Team-taught: adjusted load is divided by number of faculty members
         # Co-convened: adjusted load is divided by courses (should be 2) to add up to one full course
         courseObject.load = adjustedLoad

# Calculate the total load and percentage for each faculty member
for emplid, faculty in facultyDict.items():

   # Calculate load total (before carrer-track or tenure-track adjustement)
   faculty.calculateTotalLoad()

   # Calculate TT and CT percentage
   # Current version calculates both, final version should take ct or tt status
   # from policy and calculate either one accordingly
   ttValue, ctValue = faculty.calculatePercentage()

   

   # Print results for testing purposes
   #print(f"Faculty: {faculty.name} (ID: {faculty.emplid})")
   #print(f"  Total Workload Load: {faculty.totalLoad:.2f}%")
   #print(f"  TT Percentage (baseline 30%): {ttValue:.2f}%")
   #print(f"  CT Percentage (baseline 40%): {ctValue:.2f}%\n")

   # TO DO: include building category in course object (for composite key)
   # TO DO: adjustable calculation values importend from policy file
