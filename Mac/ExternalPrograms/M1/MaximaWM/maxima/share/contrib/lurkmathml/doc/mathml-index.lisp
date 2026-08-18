(in-package :cl-info)
(let (
(deffn-defvr-pairs '(
; CONTENT: (<INDEX TOPIC> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS> <NODE NAME>))
("mathml" . ("mathml.info" 893 3587 "Definitions for package mathml"))
("mathml_non_numeric_subscripts" . ("mathml.info" 4571 1426 "Definitions for package mathml"))
("mathml_underscore_is_subscript" . ("mathml.info" 5998 1622 "Definitions for package mathml"))
))
(section-pairs '(
; CONTENT: (<NODE NAME> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS>))
("Definitions for package mathml" . ("mathml.info" 823 6707))
)))
(load-info-hashtables (maxima::maxima-load-pathname-directory) deffn-defvr-pairs section-pairs))
