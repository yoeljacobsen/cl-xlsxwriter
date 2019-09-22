(in-package #:xlsxwriter)

(defun lxw-example (filename)
  (let* ((workbook (workbook-new filename))
         (worksheet (workbook-add-worksheet workbook "sheet1")))
    (worksheet-write-string worksheet 10 10 "Hello world" (null-pointer))
    (worksheet-write-number worksheet 11 10 42d0 (null-pointer))
    (workbook-close workbook)))

(defun test-helper (filename)
  (let* ((workbook (workbook-new filename))
         (worksheet (workbook-add-worksheet workbook "sheet1"))
         (items `(1 "two" 3.0 "four" 5/6 "six" "ML without any newline characters" ,(format nil "This~%is~%a multi-line~%Cell")))
         (f1 (workbook-add-format workbook))
         (f2 (workbook-add-format workbook)))
    (format-set-bold f1)
    (format-set-text-wrap f1)
    (format-set-align f1 :lxw-align-vertical-justify)
    (format-set-font-color f2 lxw-color-white)
    (format-set-bg-color f2 lxw-color-blue)
    (format-set-pattern f2 :lxw-pattern-solid)
    (format-set-align f2 :lxw-align-center)
    (format-set-border f2 :lxw-border-thick)
    (loop for val in items
          do (worksheet-write worksheet val :row 0 :format f1))
    (loop for val in items
          do (worksheet-write worksheet val :row 3 :format f2))
    (workbook-close workbook)))

