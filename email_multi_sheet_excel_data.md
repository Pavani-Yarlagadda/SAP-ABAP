
Report to send multiple sheets of excel in email -
```
*---- Includes
INCLUDE: zyp_excel_multi_sheets_top,
         zyp_excel_multi_sheets_scr,
         zyp_excel_multi_sheets_meth.

*---- Start of selection
START-OF-SELECTION.
  gv_obj->process_mail_data( IMPORTING eo_excel = DATA(lo_excel) ).
  gv_obj->send_email( CHANGING co_excel = lo_excel ).

```

Declarations part -
```
*---- Types declaration
TYPES: BEGIN OF ty_s_excel_content,
         row_no TYPE i,
         col_no TYPE i,
         value  TYPE string,
       END OF ty_s_excel_content,
       " Table types
       ty_t_excel_content TYPE STANDARD TABLE OF ty_s_excel_content.

*---- Data declaration
DATA: lt_excel_data_sheet1 TYPE ty_t_excel_content,
      lt_excel_data_sheet2 TYPE ty_t_excel_content,
      lt_att_hex           TYPE solix_tab.

DATA ls_excel_content   TYPE ty_s_excel_content.

DATA: lv_xstring   TYPE xstring,
      lv_filename  TYPE string,
      lv_main_text TYPE bcsy_text,
      lv_size      TYPE i,
      zip          TYPE xstring.

*DATA lc_zipper      TYPE REF TO cl_abap_zip.
DATA: send_request  TYPE REF TO cl_bcs,
      document      TYPE REF TO cl_document_bcs,
      recipient     TYPE REF TO if_recipient_bcs,
      bcs_exception TYPE REF TO cx_bcs,
      mailto        TYPE ad_smtpadr,
      sent_to_all   TYPE os_boolean.

CONSTANTS: gc_raw         TYPE char3      VALUE 'RAW',
           gc_subject     TYPE char50     VALUE 'Sales Order Details',
           gc_excel       TYPE char3      VALUE 'XLS',
           gc_sender_mail TYPE ad_smtpadr VALUE 'PAVANI.YARLAGADDA@VISTEX.COM',
           gc_success_msg TYPE char1      VALUE 'S',
           gc_info_msg    TYPE char1      VALUE 'I',
           gc_error_msg   TYPE char1      VALUE 'E',
           gc_mark_x      TYPE char1      VALUE 'X'.

CLASS lcl_main DEFINITION.
  PUBLIC SECTION.
    METHODS: process_mail_data EXPORTING eo_excel          TYPE REF TO cl_cmcb_excel_2007,
             send_email        CHANGING co_excel           TYPE REF TO cl_cmcb_excel_2007,
             itab_to_xstring   IMPORTING ir_data_ref       TYPE data
                               RETURNING VALUE(rv_xstring) TYPE xstring.
ENDCLASS.
```

Selection Screen -
```
*---- Selection Screen
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-000.
  "---- Parameters
  PARAMETERS: p_email TYPE ad_smtpadr.
SELECTION-SCREEN END OF BLOCK b1.

*---- Initialization
INITIALIZATION.
data(gv_obj) = new lcl_main( ).
```

Methods implemenation -
```
CLASS lcl_main IMPLEMENTATION.

  METHOD process_mail_data.

*&-----Get data to send in excel-----------------
    SELECT
      FROM vbak
      FIELDS vbeln,
             vbtyp,
             vkorg
      ORDER BY vbeln
      INTO TABLE @DATA(lt_vbak).

    SELECT
      FROM vbap
      FIELDS vbeln,
             posnr,
             matnr,
             vbtyp_ana
      ORDER BY vbeln
      INTO TABLE @DATA(lt_vbap).

*------ Convert table data into excel sheet structure

    "---- Pass VBAK data header structure
    lt_excel_data_sheet1 = VALUE #( row_no = 1 ( col_no = 1 value = 'VBELN' )
                                             ( col_no = 2 value = 'VBTYP' )
                                             ( col_no = 3 value = 'VKORG' ) ).

    "---- Pass VBAK table data
    lt_excel_data_sheet1 = VALUE #( BASE lt_excel_data_sheet1
                                   FOR ls_vbak IN lt_vbak INDEX INTO idx
                                   ( row_no = idx + 1  col_no = 1 value = ls_vbak-vbeln )
                                   ( row_no = idx + 1  col_no = 2 value = ls_vbak-vbtyp )
                                   ( row_no = idx + 1  col_no = 3 value = ls_vbak-vkorg ) ).

    "---- Pass VBAP data header structure
    lt_excel_data_sheet2 = VALUE #( row_no = 1 ( col_no = 1 value = 'VBELN' )
                                             ( col_no = 2 value = 'POSNR' )
                                             ( col_no = 3 value = 'MATNR' )
                                             ( col_no = 3 value = 'VBTYP_ANA' ) ).

    "---- Pass VBAP table data
    lt_excel_data_sheet2 = VALUE #( BASE lt_excel_data_sheet2
                                   FOR ls_vbap IN lt_vbap INDEX INTO idx
                                   ( row_no = idx + 1  col_no = 1 value = ls_vbap-vbeln )
                                   ( row_no = idx + 1  col_no = 2 value = ls_vbap-posnr )
                                   ( row_no = idx + 1  col_no = 3 value = ls_vbap-matnr )
                                   ( row_no = idx + 1  col_no = 4 value = ls_vbap-vbtyp_ana ) ).

**&-----Create XLSX file with the new Itab structures-----------
    CREATE OBJECT eo_excel.
    DATA(lo_dwnld) = NEW cl_cmcb_download_org_hierarchy( ).

    "------ Add VBAK data to excel -----------
    eo_excel->add_sheet( i_sheetname = 'SALES_HDR_DATA' ).

    LOOP AT lt_excel_data_sheet1 INTO ls_excel_content.
      IF ls_excel_content-row_no EQ 1.                       "For Header Record
        eo_excel->set_cell( i_data      = ls_excel_content-value
                            i_row_index = ls_excel_content-row_no
                            i_col_index = ls_excel_content-col_no
                            i_cellstyle = cl_srt_wsp_excel_2007=>c_cellstyle_header
                            i_sheetname = 'SALES_HDR_DATA ' ).
      ELSE.
        eo_excel->set_cell( i_data      = ls_excel_content-value   "For Data records
                            i_row_index = ls_excel_content-row_no
                            i_col_index = ls_excel_content-col_no
                            i_cellstyle = cl_srt_wsp_excel_2007=>c_cellstyle_normal
                            i_sheetname = 'SALES_HDR_DATA' ).
      ENDIF.
    ENDLOOP.

    "-------- Add VBAP data to excel ----------
    eo_excel->add_sheet( i_sheetname = 'SALES_ITEM_DATA' ).

    LOOP AT lt_excel_data_sheet2 INTO ls_excel_content.
      IF ls_excel_content-row_no EQ 1.                            "For Header Record
        eo_excel->set_cell( i_data      = ls_excel_content-value
                            i_row_index = ls_excel_content-row_no
                            i_col_index = ls_excel_content-col_no
                            i_cellstyle = cl_srt_wsp_excel_2007=>c_cellstyle_header
                            i_sheetname = 'SALES_ITEM_DATA' ).
      ELSE.
        eo_excel->set_cell( i_data      = ls_excel_content-value         "For Data records
                            i_row_index = ls_excel_content-row_no
                            i_col_index = ls_excel_content-col_no
                            i_cellstyle = cl_srt_wsp_excel_2007=>c_cellstyle_normal
                            i_sheetname = 'SALES_ITEM_DATA' ).
      ENDIF.
    ENDLOOP.


  ENDMETHOD.

*--- Method to send an email using BCS
  METHOD send_email.

    DATA: lv_xstring TYPE xstring.

    TRY.
        "Create send request
        DATA(lo_send_request) = cl_bcs=>create_persistent( ).
        DATA(lv_file_name)    = 'SALES_DATA.XLSX'.

        "Create mail body
        DATA(lt_body) = VALUE bcsy_text(
                          ( line = TEXT-007 ) ( )
                          ( line = TEXT-008 ) ( )
                          ( line = TEXT-009 ) ).

        "Set up document object
        DATA(lo_document) = cl_document_bcs=>create_document(
          i_type    = gc_raw
          i_text    = lt_body
          i_subject = gc_subject ).

*        GET REFERENCE OF lt_so_data INTO DATA(lo_data_ref).
*        lv_xstring = gv_obj->itab_to_xstring( co_excel ).

        lv_xstring = co_excel->transform( ).

        lo_document->add_attachment(
          i_attachment_type    = gc_excel
          i_attachment_size    = CONV #( xstrlen( lv_xstring ) )
          i_attachment_subject = gc_subject
          i_attachment_header  = VALUE #( ( line = lv_file_name ) )
          i_att_content_hex    = cl_bcs_convert=>xstring_to_solix( lv_xstring )
        ).

        "Add document to send request
        lo_send_request->set_document( lo_document ).

        "Set sender
        lo_send_request->set_sender(
          cl_cam_address_bcs=>create_internet_address(
            i_address_string = CONV #( gc_sender_mail )
          )
        ).

        "Set Recipient | This method has options to set CC/BCC as well
        lo_send_request->add_recipient(
          i_recipient = cl_cam_address_bcs=>create_internet_address(
                          i_address_string = CONV #( p_email ) )
          i_express   = abap_true ).

        "Send Email
        DATA(lv_sent_to_all) = lo_send_request->send( ).
        COMMIT WORK.

        IF lv_sent_to_all EQ gc_mark_x.
          MESSAGE TEXT-005 TYPE gc_success_msg.                           " Email sent successfully
        ELSE.
          MESSAGE TEXT-006 TYPE gc_info_msg DISPLAY LIKE gc_error_msg.    " Email not sent
        ENDIF.

      CATCH cx_send_req_bcs INTO DATA(lx_req_bsc).
      CATCH cx_document_bcs INTO DATA(lx_doc_bcs).
      CATCH cx_address_bcs  INTO DATA(lx_add_bcs).
    ENDTRY.

  ENDMETHOD.

*--- Method to convert data into xstring
  METHOD itab_to_xstring.

    FIELD-SYMBOLS: <fs_data> TYPE ANY TABLE.

    CLEAR rv_xstring.
    ASSIGN ir_data_ref->* TO <fs_data>.

    TRY.
        cl_salv_table=>factory(
          IMPORTING
            r_salv_table = DATA(lo_table)
          CHANGING
            t_table      = <fs_data> ).

        DATA(lt_fcat) =
          cl_salv_controller_metadata=>get_lvc_fieldcatalog(
          r_columns      = lo_table->get_columns( )
          r_aggregations = lo_table->get_aggregations( ) ).

        DATA(lo_result) =
          cl_salv_ex_util=>factory_result_data_table(
          r_data         = ir_data_ref
          t_fieldcatalog = lt_fcat ).

        cl_salv_bs_tt_util=>if_salv_bs_tt_util~transform(
          EXPORTING
            xml_type      = if_salv_bs_xml=>c_type_xlsx
            xml_version   = cl_salv_bs_a_xml_base=>get_version( )
            r_result_data = lo_result
            xml_flavour   = if_salv_bs_c_tt=>c_tt_xml_flavour_export
            gui_type      = if_salv_bs_xml=>c_gui_type_gui
          IMPORTING
            xml           = rv_xstring ).
      CATCH cx_root.
        CLEAR rv_xstring.
    ENDTRY.

  ENDMETHOD.

ENDCLASS.
```
