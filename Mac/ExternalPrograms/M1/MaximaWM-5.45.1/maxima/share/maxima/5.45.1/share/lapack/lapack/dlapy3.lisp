;;; Compiled by f2cl version:
;;; ("f2cl1.l,v 2edcbd958861 2012/05/30 03:34:52 toy $"
;;;  "f2cl2.l,v 96616d88fb7e 2008/02/22 22:19:34 rtoy $"
;;;  "f2cl3.l,v 96616d88fb7e 2008/02/22 22:19:34 rtoy $"
;;;  "f2cl4.l,v 96616d88fb7e 2008/02/22 22:19:34 rtoy $"
;;;  "f2cl5.l,v 3fe93de3be82 2012/05/06 02:17:14 toy $"
;;;  "f2cl6.l,v 1d5cbacbb977 2008/08/24 00:56:27 rtoy $"
;;;  "macros.l,v 3fe93de3be82 2012/05/06 02:17:14 toy $")

;;; Using Lisp CMU Common Lisp 20d (20D Unicode)
;;; 
;;; Options: ((:prune-labels nil) (:auto-save t) (:relaxed-array-decls t)
;;;           (:coerce-assigns :as-needed) (:array-type ':array)
;;;           (:array-slicing t) (:declare-common nil)
;;;           (:float-format double-float))

(in-package :lapack)


(let* ((zero 0.0))
  (declare (type (double-float 0.0 0.0) zero) (ignorable zero))
  (defun dlapy3 (x y z)
    (declare (type (double-float) z y x))
    (prog ((w 0.0) (xabs 0.0) (yabs 0.0) (zabs 0.0) (dlapy3 0.0))
      (declare (type (double-float) w xabs yabs zabs dlapy3))
      (setf xabs (abs x))
      (setf yabs (abs y))
      (setf zabs (abs z))
      (setf w (max xabs yabs zabs))
      (cond
        ((= w zero)
         (setf dlapy3 (+ xabs yabs zabs)))
        (t
         (setf dlapy3
                 (* w
                    (f2cl-lib:fsqrt
                     (+ (expt (/ xabs w) 2)
                        (expt (/ yabs w) 2)
                        (expt (/ zabs w) 2)))))))
      (go end_label)
     end_label
      (return (values dlapy3 nil nil nil)))))

(in-package #-gcl #:cl-user #+gcl "CL-USER")
#+#.(cl:if (cl:find-package '#:f2cl) '(and) '(or))
(eval-when (:load-toplevel :compile-toplevel :execute)
  (setf (gethash 'fortran-to-lisp::dlapy3
                 fortran-to-lisp::*f2cl-function-info*)
          (fortran-to-lisp::make-f2cl-finfo
           :arg-types '((double-float) (double-float) (double-float))
           :return-values '(nil nil nil)
           :calls 'nil)))

