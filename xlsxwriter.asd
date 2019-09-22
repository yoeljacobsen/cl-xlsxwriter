;;;; xlsxwriter.asd

(asdf:defsystem #:xlsxwriter
  :description "Describe xlsxwriter here"
  :author "Your Name <your.name@example.com>"
  :license "Specify license here"
  :serial t
  :depends-on (:cffi)
  :components ((:file "package")
               (:file "xlsxwriter")
               (:file "types")
               (:file "workbook")
               (:file "worksheet")
               (:file "format")
               (:file "helpers")
               (:file "tests")
               (:file "export")))

