RangeChecker
============

This software can be used to analyze data points and time ratios in/above/below a defined range using a standard linear interpolation model. Example: INR coagulation values. This method has been established by Rosendaal et al. to assess the quality of anticoagulation in patients [1]. Published data analyzed by RangeChecker: [2,3].

![Screenshot][rc_screenshot]

Features
--------

- import and validation of data
- currently a linear model is implemented, others may be added
- calculation of time and # of data points in the complete time range
- calculation for in, below and above individual range, and in, below and above safety range:
  * # of data points, % in respect to all analyzable
  * time, % in respect to all analyzable
- calculation for below and above individual range, and below and above safety range:
  * deviation from border using an area under the curve (AUC) model
- excluded are
  * time ranges longer than a cutoff limit
  * time ranges and data points with bridging (e.g. additional LMWH)

Usage
-----

The source code was built with VisualBasic Express 2010. RangeChecker offers a GUI that allows loading of one or several tab spaced files as defined below. A log file (text) and result file (tab spaced text, can be imported into a statistics program) are written into the user's main document folder.

Input data format
-----------------

current data input format: **tab spaced text** (exported from spreadsheet or database).

	@	@	CHECK	param	value
	@	@	SET	param	value
	Table legend here will be treated as comment
	#	#	UPN	min	max
	date	INR
	date	INR			(any number of data points)
	date	INR
	date	INR
	#	#	UPN	min	max
	date	INR
	date	INR	B		(comment: first day of bridging)
	date	INR	B		(any number of bridging days allowed)
	date	INR	B		(comment: last day of bridging)
	date	INR
	#	#	UPN	min	max	(comment: no data for this UPN)
	#	#	UPN	min	max	(comment: just go on with the next)
	date	INR
	date	INR

**Notes:**
- @ and # are indeed the @ and # characters used as keywords
- CHECK, SET, B are keywords
- UPN is an identifier for the patient
- date is a date in the dd.mm.yyyy format
- INR, min, max are appropriate INR values (or other data values in other settings)

current **parameters** that can be set

	CHECK	PROGRAM_ID	RC
	CHECK	MIN_VERSION	2.2
	SET	SAFETY_MIN	value	(comment: preset = 2.0)
	SET	SAFETY_MAX	value	(comment: preset = 4.5)
	SET	MAX_TIME_INTERVAL value	(comment: days; preset = 100)
	SET	DELTA_YEARS	value	(comment: preset = 30)

**see example file:** [rc_example.xls][rc_example_xls]


Contact
-------

Do not hesitate to [contact][home] me if you would like to use RangeChecker for your project.


License
-------

Copyright (c) 2005-2014 Pirmin Schmid, [MIT license][license].


References
----------

1.	 Rosendaal FR, Cannegieter SC, van der Meer FJ, Briët E. A method to determine the optimal intensity of oral anticoagulant therapy. Thrombosis and haemostasis 1993;69(3):236–239
2.	Fritschi J, Raddatz-Müller P, Schmid P, Wuillemin WA. Patient self-management of long-term oral anticoagulation in Switzerland. Swiss Med Wkly 2007;137(17-18):252-8
3.	Nagler M, Bachmann LM, Schmid P, Raddatz-Müller P, Wuillemin WA. Patient self-management of oral anticoagulation with vitamin K antagonists in everyday practice: efficacy and safety in a nationwide long-term prospective cohort study. PLOS ONE 2014;9(4):e95761

[home]:http://www.pirmin-schmid.ch
[license]:https://github.com/pirminschmid/RangeChecker/tree/master/LICENSE
[rc_screenshot]:https://github.com/pirminschmid/RangeChecker/tree/master/rc_screenshot.png
[rc_example_xls]:https://github.com/pirminschmid/RangeChecker/tree/master/rc_example.xls
