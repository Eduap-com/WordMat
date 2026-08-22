(defclass cmn-file (cl-source-file)
  ((type :initform "cmn")))

(defsystem maxima-dlsode/odepack-package :pathname ""
  :components
   ((:file "package")
    (:module "src"
      :components
      ((:cmn-file "dls001")))))

(defsystem maxima-dlsode/odepack-dlsode :pathname ""
  :depends-on ("maxima-dlsode/odepack-package")
  :components
  ((:module "src"
    :components
    ((:file "dlsode"
      :depends-on ("dstode" "xerrwd" "dintdy" "dvnorm" "dewset"))
     (:file "dintdy"
      :depends-on ("xerrwd"))
     (:file "dstode"
      :depends-on ("dvnorm" "dcfode"))
     (:file "dgbfa"
      :depends-on ("idamax" "dscal" "dvnorm"))
     (:file "dcfode")
     (:file "dprepj"
      :depends-on ("dgbfa" "dgefa" "dvnorm"))
     (:file "dsolsy"
      :depends-on ("dgbsl" "dgesl"))
     (:file "dewset")
     (:file "dvnorm")
     (:file "dsrcom")
     (:file "dgefa"
      :depends-on ("idamax" "dscal" "daxpy"))
     (:file "dgesl"
      :depends-on ("daxpy" "ddot"))
     (:file "dgbsl"
      :depends-on ("daxpy" "ddot"))
     (:file "dscal")
     (:file "ddot")
     (:file "idamax")
     (:file "dumach"
      :depends-on ("dumsum"))
     (:file "xerrwd"
      :depends-on ("ixsav"))
     (:file "xsetun"
      :depends-on ("ixsav"))
     (:file "xsetf"
      :depends-on ("ixsav"))
     (:file "ixsav")
     (:file "iumach")
     (:file "daxpy")
     (:file "dumsum")))))

(defsystem maxima-dlsode :pathname ""
  :depends-on ("maxima-dlsode/odepack-dlsode")
  :components
  ((:file "dlsode-interface")))

