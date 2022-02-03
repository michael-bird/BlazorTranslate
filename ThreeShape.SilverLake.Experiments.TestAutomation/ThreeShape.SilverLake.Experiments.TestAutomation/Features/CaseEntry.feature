Feature: Case entry

@local
Scenario: Create new case
	Given I use local website
		And I have started LabStar website in Chrome
		And proceed to the main page
	When I open Case Entry
		When fill the form
		| Field              | Value                              |
		| Doctor/Lab         | Texas Dental Clinic [Cruz, Teddy ] |
		| Patient first name | Jack                               |
		| Patient last name  | Smith                              |
		| Item type          | Crown                              |
		| Tooth              | 5                                  |
		| Item               | PFM High Noble                     |
		| Schedule           | HyperWorks                         |
		And click Create button
	Then case should be created and visible in the Manufacturing Manager

@Igor
Scenario: Create new case2
	Given I use Igor website
		And I have started LabStar website in Chrome
		And proceed to the main page
	When I open Case Entry
		When fill the form
		| Field              | Value          |
		| Doctor/Lab         | Test Client    |
		| Patient first name | Jack           |
		| Patient last name  | Smith          |
		| Item type          | Crown          |
		| Tooth              | 5              |
		| Item               | PFM High Noble |
		And click Create button
	Then case should be created and visible in the Manufacturing Manager