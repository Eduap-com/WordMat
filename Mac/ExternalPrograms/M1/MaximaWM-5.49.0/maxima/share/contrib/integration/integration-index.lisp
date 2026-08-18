(in-package :cl-info)
(let (
(deffn-defvr-pairs '(
; CONTENT: (<INDEX TOPIC> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS> <NODE NAME>))
("conditional_integrate" . ("abs_integrate.info" 10113 903 "Definitions for abs_integrate"))
("convert_to_signum" . ("abs_integrate.info" 11017 673 "Definitions for abs_integrate"))
("extra_definite_integration_methods" . ("abs_integrate.info" 6132 1135 "Definitions for abs_integrate"))
("extra_integration_methods" . ("abs_integrate.info" 3908 2223 "Definitions for abs_integrate"))
("intfudu" . ("abs_integrate.info" 7268 1550 "Definitions for abs_integrate"))
("intfugudu" . ("abs_integrate.info" 8819 834 "Definitions for abs_integrate"))
("signum_to_abs" . ("abs_integrate.info" 9654 458 "Definitions for abs_integrate"))
))
(section-pairs '(
; CONTENT: (<NODE NAME> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS>))
("Definitions for abs_integrate" . ("abs_integrate.info" 3840 7850))
("Introduction to abs_integrate" . ("abs_integrate.info" 696 2994))
)))
(load-info-hashtables (maxima::maxima-load-pathname-directory) deffn-defvr-pairs section-pairs))
