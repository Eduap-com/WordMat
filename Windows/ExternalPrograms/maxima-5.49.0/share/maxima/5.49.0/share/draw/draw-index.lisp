(in-package :cl-info)
(let (
(deffn-defvr-pairs '(
; CONTENT: (<INDEX TOPIC> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS> <NODE NAME>))
("plot_vector_field" . ("drawutils.info" 971 766 "Vector fields"))
("plot_vector_field3d" . ("drawutils.info" 1738 858 "Vector fields"))
("vennplot" . ("drawutils.info" 3098 314 "Venn diagrams"))
))
(section-pairs '(
; CONTENT: (<NODE NAME> . (<FILENAME> <BYTE OFFSET> <LENGTH IN CHARACTERS>))
)))
(load-info-hashtables (maxima::maxima-load-pathname-directory) deffn-defvr-pairs section-pairs))
