A small program to convert the Ad-hoc data into the required Exam MarksUpload format for SIS.
One Ad-hoc file may contain data for multiple courses. The output should be split by course: generate one separate Exam Marks Upload format file per course.
Once completed, the tool should be installable/runnable on a user’s computer and be simple to use. The user should be able to complete the conversion by selecting/importing the Ad-hoc file, with no additional complicated steps
Component Mapping (Add Component Names)
In each course output file, populate the required component names using this mapping:
Exam → Component 1
Attendance → Component 9
Other mark types → Components 2–8,10,11… (as applicable per data)
Course should be according the course list provided
Data of “Attendance” should be removed
Handle case of Empty field in ad-hoc file
