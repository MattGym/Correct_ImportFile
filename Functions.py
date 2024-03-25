from Functions_global import *

def update_alm_and_msg(sheet1, sheet2, am100_av, file1_col_tag, file1_col_seq,  file1_col_template, file1_col_event,
                       file1_col_alm, file1_col_msg, file1_col_prio, file1_col_hhprio, file1_col_hprio, file1_col_lprio,
                       file1_col_llprio, file1_col_ext1prio, file1_col_ext2prio, file1_col_ext3prio, file1_col_ext4prio,
                       file1_col_faprio, max_row1, max_row2):
    """
    Function 'update_alp_and_msg' update alarm and message group of core loop, according to given relationship table.
    ----------
    """
    for i in range(2, max_row1 + 1):
        if get_cell_value(sheet1, i, file1_col_seq) == 0:
            prefix = str(get_cell_value(sheet1, i+1, file1_col_tag))[0:3]
            found = False
            for j in range(2, max_row2+1):
                if found is False:
                    if str(get_cell_value(sheet1, i+1, file1_col_alm)) == str(get_cell_value(sheet2, j, 1)) \
                            and (prefix == str(get_cell_value(sheet2, j, 2)) or get_cell_value(sheet2, j, 2) is None):
                        if str(get_cell_value(sheet1, i, file1_col_template)) == 'Am100' and am100_av == 1:
                            if get_cell_value(sheet1, i, file1_col_hhprio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_hprio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_lprio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_llprio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_ext1prio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_ext2prio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_ext3prio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_ext4prio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_faprio) == 900:
                                # print(i, '-.', prefix, ' -template ', str(get_cell_value(sheet1, i, file1_col_template)), " -seq ", str(get_cell_value(sheet1, i, file1_col_seq)), " -.1-DOC ",str(get_cell_value(sheet1, i+1, file1_col_alm)), ' A1- ', str(get_cell_value(sheet1, i, file1_col_ext1prio)), ' A2- ',str(get_cell_value(sheet1, i, file1_col_ext2prio)), ' A3- ', str(get_cell_value(sheet1, i, file1_col_ext3prio)), ' - ',str(get_cell_value(sheet2, j, 2)), ' - ',str(get_cell_value(sheet2, j, 3)))
                                set_cell_value(sheet1, i, file1_col_alm, get_cell_value(sheet2, j, 3))
                                set_cell_value(sheet1, i, file1_col_msg, get_cell_value(sheet2, j, 3))
                                found = True
                            else:
                                # print(i, '-.', prefix, ' -template ', str(get_cell_value(sheet1, i, file1_col_template)), " -seq ", str(get_cell_value(sheet1, i, file1_col_seq)), " -.1-DOC ",str(get_cell_value(sheet1, i+1, file1_col_alm)), ' A1- ', str(get_cell_value(sheet1, i, file1_col_ext1prio)), ' A2- ',str(get_cell_value(sheet1, i, file1_col_ext2prio)), ' A3- ', str(get_cell_value(sheet1, i, file1_col_ext3prio)), ' - ',str(get_cell_value(sheet2, j, 2)), ' - ',str(get_cell_value(sheet2, j, 3)))
                                set_cell_value(sheet1, i, file1_col_alm, get_cell_value(sheet2, j, 4))
                                set_cell_value(sheet1, i, file1_col_msg, get_cell_value(sheet2, j, 4))
                                found = True
                        if str(get_cell_value(sheet1, i, file1_col_template))[0:4] == 'Am10' \
                                and str(get_cell_value(sheet1, i, file1_col_template)) != 'Am100':
                            if get_cell_value(sheet1, i, file1_col_hhprio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_hprio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_lprio) == 900 \
                                    or get_cell_value(sheet1, i, file1_col_llprio) == 900:
                                # print(i, '-.', prefix, ' -template ', str(get_cell_value(sheet1, i, file1_col_template)), " -seq ", str(get_cell_value(sheet1, i, file1_col_seq)), " -.1-DOC ",str(get_cell_value(sheet1, i+1, file1_col_alm)), ' A1- ', str(get_cell_value(sheet1, i, file1_col_ext1prio)), ' A2- ',str(get_cell_value(sheet1, i, file1_col_ext2prio)), ' A3- ', str(get_cell_value(sheet1, i, file1_col_ext3prio)), ' - ',str(get_cell_value(sheet2, j, 2)), ' - ',str(get_cell_value(sheet2, j, 3)))
                                set_cell_value(sheet1, i, file1_col_alm, get_cell_value(sheet2, j, 3))
                                set_cell_value(sheet1, i, file1_col_msg, get_cell_value(sheet2, j, 3))
                                found = True
                            else:
                                # print(i, '-.', prefix, ' -template ', str(get_cell_value(sheet1, i, file1_col_template)), " -seq ", str(get_cell_value(sheet1, i, file1_col_seq)), " -.1-DOC ",str(get_cell_value(sheet1, i+1, file1_col_alm)), ' A1- ', str(get_cell_value(sheet1, i, file1_col_ext1prio)), ' A2- ',str(get_cell_value(sheet1, i, file1_col_ext2prio)), ' A3- ', str(get_cell_value(sheet1, i, file1_col_ext3prio)), ' - ',str(get_cell_value(sheet2, j, 2)), ' - ',str(get_cell_value(sheet2, j, 3)))
                                set_cell_value(sheet1, i, file1_col_alm, get_cell_value(sheet2, j, 4))
                                set_cell_value(sheet1, i, file1_col_msg, get_cell_value(sheet2, j, 4))
                                found = True
                        if str(get_cell_value(sheet1, i, file1_col_template))[0:4] == 'Dm10':
                            if get_cell_value(sheet1, i, file1_col_event) == '1' \
                                    and get_cell_value(sheet1, i, file1_col_prio) == '900':
                                # print(i, '-.', prefix, ' -template ', str(get_cell_value(sheet1, i, file1_col_template)), " -seq ", str(get_cell_value(sheet1, i, file1_col_seq)), " -.1-DOC ",str(get_cell_value(sheet1, i+1, file1_col_alm)), ' - ',str(get_cell_value(sheet2, j, 2)), ' - ',str(get_cell_value(sheet2, j, 3)))

                                set_cell_value(sheet1, i, file1_col_alm, get_cell_value(sheet2, j, 3))
                                set_cell_value(sheet1, i, file1_col_msg, get_cell_value(sheet2, j, 3))
                                found = True
                            else:

                                # print(i, '-.', prefix, ' -template ', str(get_cell_value(sheet1, i, file1_col_template)), " -seq ", str(get_cell_value(sheet1, i, file1_col_seq)), " -.1-DOC ",str(get_cell_value(sheet1, i+1, file1_col_alm)), ' - ',str(get_cell_value(sheet2, j, 2)), ' - ',str(get_cell_value(sheet2, j, 3)))

                                set_cell_value(sheet1, i, file1_col_alm, get_cell_value(sheet2, j, 4))
                                set_cell_value(sheet1, i, file1_col_msg, get_cell_value(sheet2, j, 4))
                                found = True
                        if str(get_cell_value(sheet1, i, file1_col_template))[0:4] != 'Am10' \
                                and str(get_cell_value(sheet1, i, file1_col_template))[0:4] != 'Dm10':
                            # print(i, '-.', prefix, ' -template ', str(get_cell_value(sheet1, i, file1_col_template)),
                            #      " -seq ", str(get_cell_value(sheet1, i, file1_col_seq)), " -.1-DOC ",
                            #       str(get_cell_value(sheet1, i + 1, file1_col_alm)), ' - ',
                            #      str(get_cell_value(sheet2, j, 2)), ' - ', str(get_cell_value(sheet2, j, 3)))

                            set_cell_value(sheet1, i, file1_col_alm, get_cell_value(sheet2, j, 4))
                            set_cell_value(sheet1, i, file1_col_msg, get_cell_value(sheet2, j, 4))
                            found = True


def am100_clean_devicetags(sheet1, file1_col_template,  file1_col_devicetag2, file1_col_devicetag3, file1_col_devicetag4,
                           file1_col_devicetag5, file1_col_devicetag6, file1_col_devicetag7, file1_col_devicetag8,
                           file1_col_devicetag9, file1_col_devicetag10, max_rows1):
    """
    Function 'am100_clean_devicetag' cleaning devicetag 2 --> 10 for Am100 loops only and prepare to dynamic filling.
    ----------
    """""
    for i in range(2, max_rows1+1):
        if str(get_cell_value(sheet1, i, file1_col_template)) == 'Am100':
            set_cell_value(sheet1, i, file1_col_devicetag2, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag3, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag4, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag5, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag6, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag7, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag8, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag9, '', 2)
            set_cell_value(sheet1, i, file1_col_devicetag10, '', 2)


def am100_update_devicetags(sheet1, file1_col_tag, file1_col_seq, file1_col_prio, file1_col_template, file1_col_instrument_code,
                            file1_col_devicetag1, file1_col_devicetag2, file1_col_devicetag3, file1_col_devicetag4,
                            file1_col_devicetag5, file1_col_devicetag6, file1_col_devicetag7, file1_col_devicetag8,
                            file1_col_devicetag9, file1_col_devicetag10, file1_col_HHprio, file1_col_Hprio,
                            file1_col_Lprio, file1_col_LLprio, file1_col_FAprio, file1_col_Ext1prio, file1_col_Ext2prio,
                            file1_col_Ext3prio, file1_col_Ext4prio, max_rows1):
    for i in range(2, max_rows1+1):
        if str(get_cell_value(sheet1, i, file1_col_template)) == 'Am100' and get_cell_value(sheet1, i, file1_col_seq) == 0:
            tag = str(get_cell_value(sheet1, i, file1_col_tag))
            for j in range(1, 10):
                if tag == str(get_cell_value(sheet1, i+j, file1_col_tag)) and get_cell_value(sheet1, i+j, file1_col_seq) > 0:
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-2:]) == 'XA':
                        set_cell_value(sheet1, i, file1_col_FAprio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag2, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-2:]) == 'A1':
                        set_cell_value(sheet1, i, file1_col_Ext1prio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag7, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-2:]) == 'A2':
                        set_cell_value(sheet1, i, file1_col_Ext2prio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag8, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-2:]) == 'A3':
                        set_cell_value(sheet1, i, file1_col_Ext3prio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag9, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-2:]) == 'A4':
                        set_cell_value(sheet1, i, file1_col_Ext4prio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag10, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-3:]) == 'AHH':
                        set_cell_value(sheet1, i, file1_col_HHprio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag3, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-2:]) == 'AH' and str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-3:]) != 'AHH':
                        set_cell_value(sheet1, i, file1_col_Hprio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag4, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-2:]) == 'AL' and str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-3:]) != 'ALL':
                        set_cell_value(sheet1, i, file1_col_Lprio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag5, tag + '.' + str(j), 1)
                    if str(get_cell_value(sheet1, i+j, file1_col_instrument_code)[-3:]) == 'ALL':
                        set_cell_value(sheet1, i, file1_col_LLprio, get_cell_value(sheet1, i+j, file1_col_prio))
                        set_cell_value(sheet1, i, file1_col_devicetag6, tag + '.' + str(j), 1)


def merge_alarms(sheet1, file1_col_template, file1_col_seq, file1_col_HHca, file1_col_Hca, file1_col_Lca,
                 file1_col_LLca, file1_col_limit1, file1_col_limit2, file1_col_limit3, file1_col_limit4, max_rows1):
    for i in range(2, max_rows1 + 1):
        if str(get_cell_value(sheet1, i, file1_col_template)) == 'Am10' and get_cell_value(sheet1, i, file1_col_seq) == 0:
            if get_cell_value(sheet1, i, file1_col_HHca) == 'X' or int(get_cell_value(sheet1, i, file1_col_HHca)) == 1:
                set_cell_value(sheet1, i, file1_col_limit4, 'HH', 1)
            if get_cell_value(sheet1, i, file1_col_Hca) == 'X' or int(get_cell_value(sheet1, i, file1_col_Hca)) == 1:
                set_cell_value(sheet1, i, file1_col_limit3, 'H', 1)
            if get_cell_value(sheet1, i, file1_col_Lca) == 'X' or int(get_cell_value(sheet1, i, file1_col_Lca)) == 1:
                set_cell_value(sheet1, i, file1_col_limit2, 'L', 1)
            if get_cell_value(sheet1, i, file1_col_LLca) == 'X' or int(get_cell_value(sheet1, i, file1_col_LLca)) == 1:
                set_cell_value(sheet1, i, file1_col_limit1, 'LL', 1)
                



