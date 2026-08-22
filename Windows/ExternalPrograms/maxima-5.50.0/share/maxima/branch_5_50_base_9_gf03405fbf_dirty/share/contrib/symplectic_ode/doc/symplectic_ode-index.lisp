(in-package :cl-info)
(let (
(deffn-defvr-pairs '(
; CONTENT: (<INDEX TOPIC> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS> <NODE NAME>))
("poisson_bracket" . ("symplectic_ode.info" 2357 712 "Definitions for symplectic_ode"))
("symplectic_ode" . ("symplectic_ode.info" 3070 3069 "Definitions for symplectic_ode"))
))
(section-pairs '(
; CONTENT: (<NODE NAME> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS>))
("Definitions for symplectic_ode" . ("symplectic_ode.info" 2287 3852))
("Introduction to symplectic_ode" . ("symplectic_ode.info" 702 1432))
)))
(load-info-hashtables (maxima::maxima-load-pathname-directory) deffn-defvr-pairs section-pairs))
