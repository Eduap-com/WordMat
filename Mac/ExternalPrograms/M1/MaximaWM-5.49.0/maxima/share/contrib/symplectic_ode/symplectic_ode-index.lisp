(in-package :cl-info)
(let (
(deffn-defvr-pairs '(
; CONTENT: (<INDEX TOPIC> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS> <NODE NAME>))
("poisson_bracket" . ("symplectic_ode.info" 2155 712 "Definitions for symplectic_ode"))
("symplectic_ode" . ("symplectic_ode.info" 2868 2782 "Definitions for symplectic_ode"))
))
(section-pairs '(
; CONTENT: (<NODE NAME> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS>))
("Definitions for symplectic_ode" . ("symplectic_ode.info" 2085 3565))
("Introduction to symplectic_ode" . ("symplectic_ode.info" 502 1430))
)))
(load-info-hashtables (maxima::maxima-load-pathname-directory) deffn-defvr-pairs section-pairs))
