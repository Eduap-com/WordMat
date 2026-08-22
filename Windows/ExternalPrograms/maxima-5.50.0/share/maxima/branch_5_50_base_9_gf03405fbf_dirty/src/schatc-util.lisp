;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;
;;; This file is part of the Maxima computer algebra system
;;; (https://sourceforge.net/projects/maxima/)
;;;
;;; Maxima is copyrighted by its authors and licensed under the GNU
;;; General Public License.  See COPYING and AUTHORS for details.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(in-package :maxima)

;; Utility functions for working with M2 pattern matching results

(defun cdras (a b)
  "Extract the value associated with key A from association list B.
   This is commonly used to extract matched variables from M2 pattern matching results."
  (cdr (assoc a b :test #'equal)))
