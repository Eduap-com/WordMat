;;; Compiled by f2cl version:
;;; ("f2cl1.l,v 95098eb54f13 2013/04/01 00:45:16 toy $"
;;;  "f2cl2.l,v 95098eb54f13 2013/04/01 00:45:16 toy $"
;;;  "f2cl3.l,v 96616d88fb7e 2008/02/22 22:19:34 rtoy $"
;;;  "f2cl4.l,v 96616d88fb7e 2008/02/22 22:19:34 rtoy $"
;;;  "f2cl5.l,v 95098eb54f13 2013/04/01 00:45:16 toy $"
;;;  "f2cl6.l,v 1d5cbacbb977 2008/08/24 00:56:27 rtoy $"
;;;  "macros.l,v 1409c1352feb 2013/03/24 20:44:50 toy $")

;;; Using Lisp CMU Common Lisp snapshot-2020-04 (21D Unicode)
;;; 
;;; Options: ((:prune-labels nil) (:auto-save t) (:relaxed-array-decls t)
;;;           (:coerce-assigns :as-needed) (:array-type ':array)
;;;           (:array-slicing t) (:declare-common nil)
;;;           (:float-format double-float))

(in-package "HOMPACK")


(defun mulp (xxxx yyyy zzzz)
  (declare (type (array double-float (*)) zzzz yyyy xxxx))
  (f2cl-lib:with-multi-array-data
      ((xxxx double-float xxxx-%data% xxxx-%offset%)
       (yyyy double-float yyyy-%data% yyyy-%offset%)
       (zzzz double-float zzzz-%data% zzzz-%offset%))
    (prog ()
      (declare)
      (setf (f2cl-lib:fref zzzz-%data% (1) ((1 2)) zzzz-%offset%)
              (-
               (* (f2cl-lib:fref xxxx-%data% (1) ((1 2)) xxxx-%offset%)
                  (f2cl-lib:fref yyyy-%data% (1) ((1 2)) yyyy-%offset%))
               (* (f2cl-lib:fref xxxx-%data% (2) ((1 2)) xxxx-%offset%)
                  (f2cl-lib:fref yyyy-%data% (2) ((1 2)) yyyy-%offset%))))
      (setf (f2cl-lib:fref zzzz-%data% (2) ((1 2)) zzzz-%offset%)
              (+
               (* (f2cl-lib:fref xxxx-%data% (1) ((1 2)) xxxx-%offset%)
                  (f2cl-lib:fref yyyy-%data% (2) ((1 2)) yyyy-%offset%))
               (* (f2cl-lib:fref xxxx-%data% (2) ((1 2)) xxxx-%offset%)
                  (f2cl-lib:fref yyyy-%data% (1) ((1 2)) yyyy-%offset%))))
      (go end_label)
     end_label
      (return (values nil nil nil)))))

(in-package #-gcl #:cl-user #+gcl "CL-USER")
#+#.(cl:if (cl:find-package '#:f2cl) '(and) '(or))
(eval-when (:load-toplevel :compile-toplevel :execute)
  (setf (gethash 'fortran-to-lisp::mulp fortran-to-lisp::*f2cl-function-info*)
          (fortran-to-lisp::make-f2cl-finfo
           :arg-types '((array double-float (*)) (array double-float (*))
                        (array double-float (*)))
           :return-values '(nil nil nil)
           :calls 'nil)))

