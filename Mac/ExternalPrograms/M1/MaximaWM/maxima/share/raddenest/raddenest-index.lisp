(in-package :cl-info)
(let (
(deffn-defvr-pairs '(
; CONTENT: (<INDEX TOPIC> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS> <NODE NAME>))
("raddenest" . ("raddenest.info" 777 4535 "Functions and variables for package raddenest"))
))
(section-pairs '(
; CONTENT: (<NODE NAME> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS>))
("Functions and variables for package raddenest" . ("raddenest.info" 677 4635))
)))
(load-info-hashtables (maxima::maxima-load-pathname-directory) deffn-defvr-pairs section-pairs))
