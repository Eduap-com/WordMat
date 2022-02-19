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
;;;           (:float-format single-float))

(in-package "FFTPACK5")


(defun c1fm1b (n inc c ch wa fnf fac)
  (declare (type (double-float) fnf)
           (type (array double-float (*)) fac wa ch)
           (type (array f2cl-lib:complex16 (*)) c)
           (type (f2cl-lib:integer4) inc n))
  (f2cl-lib:with-multi-array-data
      ((c f2cl-lib:complex16 c-%data% c-%offset%)
       (ch double-float ch-%data% ch-%offset%)
       (wa double-float wa-%data% wa-%offset%)
       (fac double-float fac-%data% fac-%offset%))
    (prog ((nbr 0) (lid 0) (ido 0) (l2 0) (ip 0) (k1 0) (iw 0) (l1 0) (na 0)
           (nf 0) (inc2 0))
      (declare (type (f2cl-lib:integer4) inc2 nf na l1 iw k1 ip l2 ido lid
                                         nbr))
      (setf inc2 (f2cl-lib:int-add inc inc))
      (setf nf (f2cl-lib:int fnf))
      (setf na 0)
      (setf l1 1)
      (setf iw 1)
      (f2cl-lib:fdo (k1 1 (f2cl-lib:int-add k1 1))
                    ((> k1 nf) nil)
        (tagbody
          (setf ip
                  (f2cl-lib:int
                   (f2cl-lib:fref fac-%data% (k1) ((1 *)) fac-%offset%)))
          (setf l2 (f2cl-lib:int-mul ip l1))
          (setf ido (the f2cl-lib:integer4 (truncate n l2)))
          (setf lid (f2cl-lib:int-mul l1 ido))
          (setf nbr
                  (f2cl-lib:int-add 1
                                    na
                                    (f2cl-lib:int-mul 2
                                                      (min
                                                       (the f2cl-lib:integer4
                                                            (f2cl-lib:int-sub
                                                             ip
                                                             2))
                                                       (the f2cl-lib:integer4
                                                            4)))))
          (f2cl-lib:computed-goto
           (label52 label62 label53 label63 label54 label64 label55 label65
            label56 label66)
           nbr)
         label52
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f2kb ido l1 na %copy-c inc2 ch 2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label62
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f2kb ido l1 na ch 2 %copy-c inc2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label53
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f3kb ido l1 na %copy-c inc2 ch 2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label63
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f3kb ido l1 na ch 2 %copy-c inc2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label54
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f4kb ido l1 na %copy-c inc2 ch 2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label64
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f4kb ido l1 na ch 2 %copy-c inc2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label55
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f5kb ido l1 na %copy-c inc2 ch 2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label65
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1f5kb ido l1 na ch 2 %copy-c inc2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label56
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1fgkb ido ip l1 lid na %copy-c %copy-c inc2 ch ch 2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
          (go label120)
         label66
          (let ((%copy-c
                 (f2cl-lib:make-compatible-seq (array double-float (*))
                                               c
                                               (array f2cl-lib:complex16 (*)))))
            (f2cl-lib:f2cl-copy-seq %copy-c c double-float f2cl-lib:complex16)
            (c1fgkb ido ip l1 lid na ch ch 2 %copy-c %copy-c inc2
             (f2cl-lib:array-slice wa-%data%
                                   double-float
                                   (iw)
                                   ((1 *))
                                   wa-%offset%))
            (f2cl-lib:f2cl-copy-seq c %copy-c f2cl-lib:complex16 double-float))
         label120
          (setf l1 l2)
          (setf iw
                  (f2cl-lib:int-add iw
                                    (f2cl-lib:int-mul (f2cl-lib:int-sub ip 1)
                                                      (f2cl-lib:int-add ido
                                                                        ido))))
          (if (<= ip 5) (setf na (f2cl-lib:int-sub 1 na)))
         label125))
      (go end_label)
     end_label
      (return (values nil nil nil nil nil nil nil)))))

(in-package #-gcl #:cl-user #+gcl "CL-USER")
#+#.(cl:if (cl:find-package '#:f2cl) '(and) '(or))
(eval-when (:load-toplevel :compile-toplevel :execute)
  (setf (gethash 'fortran-to-lisp::c1fm1b
                 fortran-to-lisp::*f2cl-function-info*)
          (fortran-to-lisp::make-f2cl-finfo
           :arg-types '((fortran-to-lisp::integer4) (fortran-to-lisp::integer4)
                        (array fortran-to-lisp::complex16 (*))
                        (array double-float (*)) (array double-float (*))
                        (double-float) (array double-float (*)))
           :return-values '(nil nil nil nil nil nil nil)
           :calls '(fortran-to-lisp::c1fgkb fortran-to-lisp::c1f5kb
                    fortran-to-lisp::c1f4kb fortran-to-lisp::c1f3kb
                    fortran-to-lisp::c1f2kb))))

