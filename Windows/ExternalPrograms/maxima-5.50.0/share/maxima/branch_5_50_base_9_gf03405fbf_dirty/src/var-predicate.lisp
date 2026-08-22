;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;
;;; This file is part of the Maxima computer algebra system
;;; (https://sourceforge.net/projects/maxima/)
;;;
;;; Maxima is copyrighted by its authors and licensed under the GNU
;;; General Public License.  See COPYING and AUTHORS for details.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(in-package :maxima)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Predicate functions

;; Note: varp and freevarp are not used in this file anymore.  But
;; they are used in other files.  Someday, varp and freevarp should be
;; moved elsewhere.
(declaim (inline varp))
(defun varp (x)
  (declare (special var))
  (alike1 x var))

(defun freevar (a)
  (declare (special var))
  (cond ((atom a) (not (eq a var)))
	((varp a) nil)
	((and (not (atom (car a)))
	      (member 'array (cdar a) :test #'eq))
	 (cond ((freevar (cdr a)) t)
	       (t (merror "~&FREEVAR: variable of integration appeared in subscript."))))
	(t (and (freevar (car a)) (freevar (cdr a))))))

;; Same as varp, but the second arg specifies the variable to be
;; tested instead of using the special variable VAR.
(defun varp2 (x var2)
  (alike1 x var2))

;; Like freevar but the second arg specifies the variable to be tested
;; instead of using the special variable VAR.
(defun freevar2 (a var2)
  (cond ((atom a) (not (eq a var2)))
	((varp2 a var2) nil)
	((and (not (atom (car a)))
	      (member 'array (cdar a) :test #'eq))
	 (cond ((freevar2 (cdr a) var2) t)
	       (t (merror "~&FREEVAR: variable of integration appeared in subscript."))))
	(t (and (freevar2 (car a) var2) (freevar2 (cdr a) var2)))))

