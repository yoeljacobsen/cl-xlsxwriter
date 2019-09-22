(in-package #:xlsxwriter)

(defcfun "workbook_new" lxw-workbook (filename :string))

(defcfun "workbook_add_worksheet" lxw-worksheet
  (workbook lxw-workbook)
  (sheet-name :string))

(defcfun "workbook_close" lxw-error
  (workbook lxw-workbook))

(defcfun "workbook_add_format" lxw-format
  (workbook lxw-workbook))

