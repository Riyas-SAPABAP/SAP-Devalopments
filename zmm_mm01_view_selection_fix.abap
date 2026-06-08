REPORT zmm_mm01_view_selection_fix.

*---------------------------------------------------------------------*
* Compile-safe MM01 BDC view-selection replacement.
*
* Use the FORM below to replace only the 0070 view-selection section in
* ZMM_MM01_CREATE_BDC->FORM BUILD_BDC. It preserves the same business
* views from the original program; only the table-control row selection
* is corrected so later MM01 screens align with the BDC fields.
*---------------------------------------------------------------------*

CONSTANTS:
  c_x         TYPE c VALUE 'X',
  c_ok_p_plus TYPE bdcdata-fval VALUE '=P+',
  c_ok_schl   TYPE bdcdata-fval VALUE '=SCHL',
  c_prog_mm   LIKE bdcdata-program VALUE 'SAPLMGMM',
  c_scr_0070  LIKE bdcdata-dynpro  VALUE '0070'.

DATA:
  gt_bdcdata TYPE STANDARD TABLE OF bdcdata,
  gs_bdcdata TYPE bdcdata.

START-OF-SELECTION.
  PERFORM mm01_select_views_corrected.

*---------------------------------------------------------------------*
* Corrected MM01 View Selection
*---------------------------------------------------------------------*
FORM mm01_select_views_corrected.

  "===============================================================
  " 0070 - View Selection PAGE 1
  "
  " Visible rows 01-10:
  "  01 Basic Data 1             YES
  "  02 Basic Data 2             YES
  "  03 Classification           skip
  "  04 Sales: Org Data 1        YES
  "  05 Sales: Org Data 2        YES
  "  06 Sales: General/Plant     YES
  "  07 Extended SPP Basic Data  skip
  "  08 Int'l Trade: Export      YES
  "  09 Sales Text               skip
  "  10 Purchasing               YES
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0070.
  PERFORM bdc_field  USING 'BDC_CURSOR'           'MSICHTAUSW-DYTXT(10)'.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(01)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(02)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(05)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(06)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(08)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(10)' c_x.
  PERFORM bdc_field  USING 'BDC_OKCODE'           c_ok_p_plus.

  "===============================================================
  " 0070 - View Selection PAGE 2
  "
  " After the first =P+, rows 11 and 12 are still before MRP 1:
  "  01 Int'l Trade: Import      skip
  "  02 Purchase Order Text      skip
  "  03 MRP 1                    YES
  "  04 MRP 2                    YES
  "  05 MRP 3                    YES
  "  06 MRP 4                    YES
  "  07 Advanced Planning        skip
  "  08 Extended SPP             skip
  "  09 Forecasting              skip
  "  10 Gen Plant/Storage 1      YES
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0070.
  PERFORM bdc_field  USING 'BDC_CURSOR'           'MSICHTAUSW-DYTXT(10)'.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(03)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(05)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(06)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(10)' c_x.
  PERFORM bdc_field  USING 'BDC_OKCODE'           c_ok_p_plus.

  "===============================================================
  " 0070 - View Selection PAGE 3
  "
  " After the second =P+, the remaining required views are visible:
  "  01 Gen Plant/Storage 2      YES
  "  02 Warehouse Mgmt 1         skip
  "  03 Warehouse Mgmt 2         skip
  "  04 Quality Management       YES
  "  05 Accounting 1             YES
  "  06 Accounting 2             YES
  "  07 Costing 1                YES
  "  08 Costing 2                YES
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0070.
  PERFORM bdc_field  USING 'BDC_CURSOR'           'MSICHTAUSW-DYTXT(08)'.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(01)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(05)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(06)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(07)' c_x.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(08)' c_x.
  PERFORM bdc_field  USING 'BDC_OKCODE'           c_ok_schl.

ENDFORM.

*---------------------------------------------------------------------*
* Minimal BDC helpers included only so this correction source syntax-
* checks independently. In the original report, keep using the existing
* BDC_DYNPRO and BDC_FIELD forms.
*---------------------------------------------------------------------*
FORM bdc_dynpro USING pv_program LIKE bdcdata-program
                      pv_dynpro  LIKE bdcdata-dynpro.

  CLEAR gs_bdcdata.
  gs_bdcdata-program  = pv_program.
  gs_bdcdata-dynpro   = pv_dynpro.
  gs_bdcdata-dynbegin = c_x.
  APPEND gs_bdcdata TO gt_bdcdata.

ENDFORM.

FORM bdc_field USING pv_fnam LIKE bdcdata-fnam
                     pv_fval TYPE any.

  CLEAR gs_bdcdata.
  gs_bdcdata-fnam = pv_fnam.
  gs_bdcdata-fval = pv_fval.

  IF gs_bdcdata-fval IS NOT INITIAL
     OR pv_fnam = 'BDC_OKCODE'
     OR pv_fnam = 'BDC_CURSOR'.
    APPEND gs_bdcdata TO gt_bdcdata.
  ENDIF.

ENDFORM.
