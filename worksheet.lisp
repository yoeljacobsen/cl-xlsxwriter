(in-package #:xlsxwriter)

(defcfun "worksheet_write_string" lxw-error
  (worksheet lxw-worksheet)
  (row lxw-row)
  (col lxw-col)
  (string :string)
  (format lxw-format))

(defcfun "worksheet_write_number" lxw-error
  (worksheet lxw-worksheet)
  (row lxw-row)
  (col lxw-col)
  (double :double)
  (format lxw-format))

(defcfun "worksheet_set_column" lxw-error
  (worksheet lxw-worksheet)
  (first-col lxw-col)
  (last-col lxw-col)
  (width :double)
  (format lxw-format))

(defcfun "worksheet_merge_range" lxw-error
  (worksheet lxw-worksheet)
  (first-row lxw-row)
  (first-col lxw-col)
  (last-row lxw-row)
  (last-col lxw-col)
  (string :string)
  (format lxw-format))

(defcfun "worksheet_insert_image" lxw-error
  (worksheet lxw-worksheet)
  (row lxw-row)
  (col lxw-col)
  (filename :string))

