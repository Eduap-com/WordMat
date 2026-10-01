(in-package :maxima)

;; Replaces Maxima's interactive questions with automatic answers.
;; Sign questions are answered "positive" (or "negative" when positive is not
;; an allowed answer), yes/no questions are answered "no".
;; A notice about the assumption is printed. Other questions give an error.
(defun retrieve (msg flag)
  (declare (ignore flag))
  (let ((obj  (and (consp msg) (third msg)))
        (kind (and (consp msg) (fourth msg))))
    (cond
      ;; asksign: "Is a positive, negative or zero?" and variants
      ((member kind '(" positive, negative or zero?" " positive or zero?"
                      " positive or negative?" " zero or nonzero?")
               :test #'equal)
       (mtell "Note: Maxima assumed that ~M is positive.~%" obj)
       '$pos)
      ((equal kind " negative or zero?")
       (mtell "Note: Maxima assumed that ~M is negative.~%" obj)
       '$neg)
      ;; askequal: "Is a equal to b?"
      ((equal kind (intl:gettext " equal to "))
       (mtell "Note: Maxima assumed that ~M is not equal to ~M.~%" obj (fifth msg))
       '$no)
      ;; askinteger etc.: "Is n an integer?"
      ((member kind '(" a " " an ") :test #'equal)
       (mtell "Note: Maxima assumed that ~M is not~A~A.~%"
              obj kind (stripdollar (fifth msg)))
       '$no)
      (t
       (merror "Maxima needs to ask: ~M" msg)))))
