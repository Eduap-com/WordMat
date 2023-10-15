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


(let* ((one 1.0d0) (zero 0.0d0))
  (declare (type (double-float 1.0d0 1.0d0) one)
           (type (double-float 0.0d0 0.0d0) zero)
           (ignorable one zero))
  (defun zlanhe (norm uplo n a lda work)
    (declare (type (array double-float (*)) work)
             (type (array f2cl-lib:complex16 (*)) a)
             (type (f2cl-lib:integer4) lda n)
             (type (string *) uplo norm))
    (f2cl-lib:with-multi-array-data
        ((norm character norm-%data% norm-%offset%)
         (uplo character uplo-%data% uplo-%offset%)
         (a f2cl-lib:complex16 a-%data% a-%offset%)
         (work double-float work-%data% work-%offset%))
      (prog ((absa 0.0d0) (scale 0.0d0) (sum 0.0d0) (value 0.0d0) (i 0) (j 0)
             (zlanhe 0.0d0))
        (declare (type (f2cl-lib:integer4) i j)
                 (type (double-float) absa scale sum value zlanhe))
        (cond
          ((= n 0)
           (setf value zero))
          ((lsame norm "M")
           (setf value zero)
           (cond
             ((lsame uplo "U")
              (f2cl-lib:fdo (j 1 (f2cl-lib:int-add j 1))
                            ((> j n) nil)
                (tagbody
                  (f2cl-lib:fdo (i 1 (f2cl-lib:int-add i 1))
                                ((> i
                                    (f2cl-lib:int-add j (f2cl-lib:int-sub 1)))
                                 nil)
                    (tagbody
                      (setf value
                              (max value
                                   (abs
                                    (f2cl-lib:fref a-%data%
                                                   (i j)
                                                   ((1 lda) (1 *))
                                                   a-%offset%))))
                     label10))
                  (setf value
                          (max value
                               (abs
                                (f2cl-lib:dble
                                 (f2cl-lib:fref a-%data%
                                                (j j)
                                                ((1 lda) (1 *))
                                                a-%offset%)))))
                 label20)))
             (t
              (f2cl-lib:fdo (j 1 (f2cl-lib:int-add j 1))
                            ((> j n) nil)
                (tagbody
                  (setf value
                          (max value
                               (abs
                                (f2cl-lib:dble
                                 (f2cl-lib:fref a-%data%
                                                (j j)
                                                ((1 lda) (1 *))
                                                a-%offset%)))))
                  (f2cl-lib:fdo (i (f2cl-lib:int-add j 1)
                                 (f2cl-lib:int-add i 1))
                                ((> i n) nil)
                    (tagbody
                      (setf value
                              (max value
                                   (abs
                                    (f2cl-lib:fref a-%data%
                                                   (i j)
                                                   ((1 lda) (1 *))
                                                   a-%offset%))))
                     label30))
                 label40)))))
          ((or (lsame norm "I") (lsame norm "O") (f2cl-lib:fstring-= norm "1"))
           (setf value zero)
           (cond
             ((lsame uplo "U")
              (f2cl-lib:fdo (j 1 (f2cl-lib:int-add j 1))
                            ((> j n) nil)
                (tagbody
                  (setf sum zero)
                  (f2cl-lib:fdo (i 1 (f2cl-lib:int-add i 1))
                                ((> i
                                    (f2cl-lib:int-add j (f2cl-lib:int-sub 1)))
                                 nil)
                    (tagbody
                      (setf absa
                              (abs
                               (f2cl-lib:fref a-%data%
                                              (i j)
                                              ((1 lda) (1 *))
                                              a-%offset%)))
                      (setf sum (+ sum absa))
                      (setf (f2cl-lib:fref work-%data%
                                           (i)
                                           ((1 *))
                                           work-%offset%)
                              (+
                               (f2cl-lib:fref work-%data%
                                              (i)
                                              ((1 *))
                                              work-%offset%)
                               absa))
                     label50))
                  (setf (f2cl-lib:fref work-%data% (j) ((1 *)) work-%offset%)
                          (+ sum
                             (abs
                              (f2cl-lib:dble
                               (f2cl-lib:fref a-%data%
                                              (j j)
                                              ((1 lda) (1 *))
                                              a-%offset%)))))
                 label60))
              (f2cl-lib:fdo (i 1 (f2cl-lib:int-add i 1))
                            ((> i n) nil)
                (tagbody
                  (setf value
                          (max value
                               (f2cl-lib:fref work-%data%
                                              (i)
                                              ((1 *))
                                              work-%offset%)))
                 label70)))
             (t
              (f2cl-lib:fdo (i 1 (f2cl-lib:int-add i 1))
                            ((> i n) nil)
                (tagbody
                  (setf (f2cl-lib:fref work-%data% (i) ((1 *)) work-%offset%)
                          zero)
                 label80))
              (f2cl-lib:fdo (j 1 (f2cl-lib:int-add j 1))
                            ((> j n) nil)
                (tagbody
                  (setf sum
                          (+
                           (f2cl-lib:fref work-%data%
                                          (j)
                                          ((1 *))
                                          work-%offset%)
                           (abs
                            (f2cl-lib:dble
                             (f2cl-lib:fref a-%data%
                                            (j j)
                                            ((1 lda) (1 *))
                                            a-%offset%)))))
                  (f2cl-lib:fdo (i (f2cl-lib:int-add j 1)
                                 (f2cl-lib:int-add i 1))
                                ((> i n) nil)
                    (tagbody
                      (setf absa
                              (abs
                               (f2cl-lib:fref a-%data%
                                              (i j)
                                              ((1 lda) (1 *))
                                              a-%offset%)))
                      (setf sum (+ sum absa))
                      (setf (f2cl-lib:fref work-%data%
                                           (i)
                                           ((1 *))
                                           work-%offset%)
                              (+
                               (f2cl-lib:fref work-%data%
                                              (i)
                                              ((1 *))
                                              work-%offset%)
                               absa))
                     label90))
                  (setf value (max value sum))
                 label100)))))
          ((or (lsame norm "F") (lsame norm "E"))
           (setf scale zero)
           (setf sum one)
           (cond
             ((lsame uplo "U")
              (f2cl-lib:fdo (j 2 (f2cl-lib:int-add j 1))
                            ((> j n) nil)
                (tagbody
                  (multiple-value-bind (var-0 var-1 var-2 var-3 var-4)
                      (zlassq (f2cl-lib:int-sub j 1)
                       (f2cl-lib:array-slice a-%data%
                                             f2cl-lib:complex16
                                             (1 j)
                                             ((1 lda) (1 *))
                                             a-%offset%)
                       1 scale sum)
                    (declare (ignore var-0 var-1 var-2))
                    (setf scale var-3)
                    (setf sum var-4))
                 label110)))
             (t
              (f2cl-lib:fdo (j 1 (f2cl-lib:int-add j 1))
                            ((> j (f2cl-lib:int-add n (f2cl-lib:int-sub 1)))
                             nil)
                (tagbody
                  (multiple-value-bind (var-0 var-1 var-2 var-3 var-4)
                      (zlassq (f2cl-lib:int-sub n j)
                       (f2cl-lib:array-slice a-%data%
                                             f2cl-lib:complex16
                                             ((+ j 1) j)
                                             ((1 lda) (1 *))
                                             a-%offset%)
                       1 scale sum)
                    (declare (ignore var-0 var-1 var-2))
                    (setf scale var-3)
                    (setf sum var-4))
                 label120))))
           (setf sum (* 2 sum))
           (f2cl-lib:fdo (i 1 (f2cl-lib:int-add i 1))
                         ((> i n) nil)
             (tagbody
               (cond
                 ((/= (f2cl-lib:dble (f2cl-lib:fref a (i i) ((1 lda) (1 *))))
                      zero)
                  (setf absa
                          (abs
                           (f2cl-lib:dble
                            (f2cl-lib:fref a-%data%
                                           (i i)
                                           ((1 lda) (1 *))
                                           a-%offset%))))
                  (cond
                    ((< scale absa)
                     (setf sum (+ one (* sum (expt (/ scale absa) 2))))
                     (setf scale absa))
                    (t
                     (setf sum (+ sum (expt (/ absa scale) 2)))))))
              label130))
           (setf value (* scale (f2cl-lib:fsqrt sum)))))
        (setf zlanhe value)
        (go end_label)
       end_label
        (return (values zlanhe nil nil nil nil nil nil))))))

(in-package #-gcl #:cl-user #+gcl "CL-USER")
#+#.(cl:if (cl:find-package '#:f2cl) '(and) '(or))
(eval-when (:load-toplevel :compile-toplevel :execute)
  (setf (gethash 'fortran-to-lisp::zlanhe
                 fortran-to-lisp::*f2cl-function-info*)
          (fortran-to-lisp::make-f2cl-finfo
           :arg-types '((string) (string) (fortran-to-lisp::integer4)
                        (array fortran-to-lisp::complex16 (*))
                        (fortran-to-lisp::integer4) (array double-float (*)))
           :return-values '(nil nil nil nil nil nil)
           :calls '(fortran-to-lisp::zlassq fortran-to-lisp::lsame))))

