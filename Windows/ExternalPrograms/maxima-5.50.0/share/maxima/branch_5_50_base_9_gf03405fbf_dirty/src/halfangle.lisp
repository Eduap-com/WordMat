;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;
;;; This file is part of the Maxima computer algebra project
;;; (https://sourceforge.net/projects/maxima/) 
;;; SPDX-License-Identifier: GPL-2.0-or-later 
;;;
;;; Maxima is copyrighted by its authors and licensed under the GNU
;;; General Public License.  This program is distributed WITHOUT ANY
;;; WARRANTY. See COPYING and AUTHORS for details.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; Half-angle simplification, called from trigi and trigo.  Split out
;;; from logarc so that the half-angle code's dependence on rpart's
;;; $realpart / $imagpart does not create a cycle with rpart's call to
;;; logarc.

(in-package :maxima)
(macsyma-module halfangle)

(defun halfangle (f a)
  (and (mtimesp a)
       (ratnump (cadr a))
       (equal (caddr (cadr a)) 2)
       (halfangleaux f (mul 2 a))))

(defun halfangleaux (f a) ;; f=function; a=twice argument
  (let ((sw (member f '(%cos %cot %coth %cosh) :test #'eq)))
    (cond ((member f '(%sin %cos) :test #'eq)
           (mul (halfangleaux-factor f a)
                (power (div (add 1 (porm sw (take '(%cos) a))) 2) 1//2)))
          ((member f '(%tan %cot) :test #'eq)
           (div (add 1 (porm sw (take '(%cos) a))) (take '(%sin) a)))
          ((member f '(%sinh %cosh) :test #'eq)
           (mul (halfangleaux-factor f a)
                (power (div (add (take '(%cosh) a) (porm sw 1)) 2) 1//2)))
	  ((member f '(%tanh %coth) :test #'eq)
	   (div (add (take '(%cosh) a) (porm sw 1)) (take '(%sinh) a)))
	  ((member f '(%sec %csc %sech %csch) :test #'eq)
	   (inv (halfangleaux (get f 'recip) a))))))

(defun halfangleaux-factor (f a)
  (cond 
    ((member f '(%sin %cos))
     (let ((arg (div (if (eq f '%sin) 
                         ($realpart a)
                         (add ($realpart a) '$%pi)) 
                     (mul 2 '$%pi))))
       (mul 
         (power -1 (simplify (list '($floor) arg)))
         (sub 1
           (mul 
             (add 1
               (power -1 (add (simplify (list '($floor) arg))
                              (simplify (list '($floor) (mul -1 arg))))))
               (simplify (list '($unit_step) (mul -1 ($imagpart a)))))))))
    ((member f '(%sinh %cosh))
     (let ((arg (div (add ($imagpart a) '$%pi) (mul 2 '$%pi)))
           (fac (if (eq f '%sinh)
                    (div (power (power a 2) (div 1 2)) a)
                    1)))
       (mul fac
         (power -1 (simplify (list '($floor) arg)))
         (sub 1
           (mul 
             (add 1
               (power -1 (add (simplify (list '($floor) arg))
                              (simplify (list '($floor) (mul -1 arg))))))
               (simplify (list '($unit_step) ($realpart a))))))))
    (t 1)))
