;;; Compiled by f2cl version:
;;; ("f2cl1.l,v 95098eb54f13 2013/04/01 00:45:16 toy $"
;;;  "f2cl2.l,v 95098eb54f13 2013/04/01 00:45:16 toy $"
;;;  "f2cl3.l,v 96616d88fb7e 2008/02/22 22:19:34 rtoy $"
;;;  "f2cl4.l,v 96616d88fb7e 2008/02/22 22:19:34 rtoy $"
;;;  "f2cl5.l,v 95098eb54f13 2013/04/01 00:45:16 toy $"
;;;  "f2cl6.l,v 1d5cbacbb977 2008/08/24 00:56:27 rtoy $"
;;;  "macros.l,v 1409c1352feb 2013/03/24 20:44:50 toy $")

;;; Using Lisp CMU Common Lisp snapshot-2013-11 (20E Unicode)
;;; 
;;; Options: ((:prune-labels nil) (:auto-save t) (:relaxed-array-decls t)
;;;           (:coerce-assigns :as-needed) (:array-type ':array)
;;;           (:array-slicing t) (:declare-common nil)
;;;           (:float-format single-float))

(in-package "LAPACK")


(let* ((one (f2cl-lib:cmplx 1.0d0 0.0d0)) (zero (f2cl-lib:cmplx 0.0d0 0.0d0)))
  (declare (type (f2cl-lib:complex16) one)
           (type (f2cl-lib:complex16) zero)
           (ignorable one zero))
  (defun zung2l (m n k a lda tau work info)
    (declare (type (array f2cl-lib:complex16 (*)) work tau a)
             (type (f2cl-lib:integer4) info lda k n m))
    (f2cl-lib:with-multi-array-data
        ((a f2cl-lib:complex16 a-%data% a-%offset%)
         (tau f2cl-lib:complex16 tau-%data% tau-%offset%)
         (work f2cl-lib:complex16 work-%data% work-%offset%))
      (prog ((i 0) (ii 0) (j 0) (l 0))
        (declare (type (f2cl-lib:integer4) i ii j l))
        (setf info 0)
        (cond
          ((< m 0)
           (setf info -1))
          ((or (< n 0) (> n m))
           (setf info -2))
          ((or (< k 0) (> k n))
           (setf info -3))
          ((< lda (max (the f2cl-lib:integer4 1) (the f2cl-lib:integer4 m)))
           (setf info -5)))
        (cond
          ((/= info 0)
           (xerbla "ZUNG2L" (f2cl-lib:int-sub info))
           (go end_label)))
        (if (<= n 0) (go end_label))
        (f2cl-lib:fdo (j 1 (f2cl-lib:int-add j 1))
                      ((> j (f2cl-lib:int-add n (f2cl-lib:int-sub k))) nil)
          (tagbody
            (f2cl-lib:fdo (l 1 (f2cl-lib:int-add l 1))
                          ((> l m) nil)
              (tagbody
                (setf (f2cl-lib:fref a-%data% (l j) ((1 lda) (1 *)) a-%offset%)
                        zero)
               label10))
            (setf (f2cl-lib:fref a-%data%
                                 ((f2cl-lib:int-add (f2cl-lib:int-sub m n) j)
                                  j)
                                 ((1 lda) (1 *))
                                 a-%offset%)
                    one)
           label20))
        (f2cl-lib:fdo (i 1 (f2cl-lib:int-add i 1))
                      ((> i k) nil)
          (tagbody
            (setf ii (f2cl-lib:int-add (f2cl-lib:int-sub n k) i))
            (setf (f2cl-lib:fref a-%data%
                                 ((f2cl-lib:int-add (f2cl-lib:int-sub m n) ii)
                                  ii)
                                 ((1 lda) (1 *))
                                 a-%offset%)
                    one)
            (multiple-value-bind
                  (var-0 var-1 var-2 var-3 var-4 var-5 var-6 var-7 var-8)
                (zlarf "Left" (f2cl-lib:int-add (f2cl-lib:int-sub m n) ii)
                 (f2cl-lib:int-sub ii 1)
                 (f2cl-lib:array-slice a-%data%
                                       f2cl-lib:complex16
                                       (1 ii)
                                       ((1 lda) (1 *))
                                       a-%offset%)
                 1 (f2cl-lib:fref tau-%data% (i) ((1 *)) tau-%offset%) a lda
                 work)
              (declare (ignore var-0 var-1 var-2 var-3 var-4 var-5 var-6
                               var-8))
              (setf lda var-7))
            (zscal
             (f2cl-lib:int-sub (f2cl-lib:int-add (f2cl-lib:int-sub m n) ii) 1)
             (- (f2cl-lib:fref tau-%data% (i) ((1 *)) tau-%offset%))
             (f2cl-lib:array-slice a-%data%
                                   f2cl-lib:complex16
                                   (1 ii)
                                   ((1 lda) (1 *))
                                   a-%offset%)
             1)
            (setf (f2cl-lib:fref a-%data%
                                 ((f2cl-lib:int-add (f2cl-lib:int-sub m n) ii)
                                  ii)
                                 ((1 lda) (1 *))
                                 a-%offset%)
                    (- one
                       (f2cl-lib:fref tau-%data% (i) ((1 *)) tau-%offset%)))
            (f2cl-lib:fdo (l (f2cl-lib:int-add m (f2cl-lib:int-sub n) ii 1)
                           (f2cl-lib:int-add l 1))
                          ((> l m) nil)
              (tagbody
                (setf (f2cl-lib:fref a-%data%
                                     (l ii)
                                     ((1 lda) (1 *))
                                     a-%offset%)
                        zero)
               label30))
           label40))
        (go end_label)
       end_label
        (return (values nil nil nil nil lda nil nil info))))))

(in-package #-gcl #:cl-user #+gcl "CL-USER")
#+#.(cl:if (cl:find-package '#:f2cl) '(and) '(or))
(eval-when (:load-toplevel :compile-toplevel :execute)
  (setf (gethash 'fortran-to-lisp::zung2l
                 fortran-to-lisp::*f2cl-function-info*)
          (fortran-to-lisp::make-f2cl-finfo
           :arg-types '((fortran-to-lisp::integer4) (fortran-to-lisp::integer4)
                        (fortran-to-lisp::integer4)
                        (array fortran-to-lisp::complex16 (*))
                        (fortran-to-lisp::integer4)
                        (array fortran-to-lisp::complex16 (*))
                        (array fortran-to-lisp::complex16 (*))
                        (fortran-to-lisp::integer4))
           :return-values '(nil nil nil nil fortran-to-lisp::lda nil nil
                            fortran-to-lisp::info)
           :calls '(fortran-to-lisp::zscal fortran-to-lisp::zlarf
                    fortran-to-lisp::xerbla))))

