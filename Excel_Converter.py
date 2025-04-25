import csv


def convertG10(file,opt=''):
    global ds, bd, x1, g
    f = open(file)
    if opt=='':
        g = open(file[:-4]+'.csv', 'w+', newline='')
    else:
        ds = open(file[:-4] + '(DS).csv', 'w+', newline='')
        bd = open(file[:-4] + '(BD).csv', 'w+', newline='')
        lst = open(opt)
        x1 = lst.readlines()
    x = f.readlines()

    c_name = ['ROLL_NO', 'NAME', 'ENG', 'TAMIL', 'HINDI', 'FRENCH', 'MATHS', 'SCIENCE', 'SOCIAL', 'TOTAL', 'RESULT']
    lst = []
    l2 = []
    p_f = []
    totalmark = []
    avg=[]
    final = []

    for i in range(len(x)):
        y = x[i].split()
        if y == []:
            pass
        else:
            roll = ''
            name = ''
            total = 0
            marks_lst = []
            pass_fail = []
            total_lst = []

            for j in y:
                name_rollno = [roll, name[1:]]
                if y[0].isdigit() and j.isalpha() and j != ('PASS' or 'FAIL'):
                    name = name + j + ' '
                elif y[0].isdigit() and j.isdigit() and len(j) > 4:
                    roll += j

            if len(name_rollno[0]) > 1 and len(name_rollno[1]) > 1:
                lst.append(name_rollno)

            if y[0].isdigit() and len(y[0])>4:
                sub_codes = []
                for code in y:
                    if code.isdigit() and len(code)==3:
                        sub_codes.append(code)

                if sub_codes[1]=='006':
                    y1=x[i + 1].split()
                    for k in range(len(y1)):
                        if k%2==0:
                            marks_lst.append(y1[k])
                    marks_lst.insert(2,'-')
                    marks_lst.insert(3,'-')
                elif sub_codes[1]=='085':
                    y1 = x[i + 1].split()
                    for k in range(len(y1)):
                        if k % 2 == 0:
                            marks_lst.append(y1[k])
                    marks_lst.insert(1, '-')
                    marks_lst.insert(3, '-')
                elif sub_codes[1]=='018':
                    y1 = x[i + 1].split()
                    for k in range(len(y1)):
                        if k % 2 == 0:
                            marks_lst.append(y1[k])
                    marks_lst.insert(1, '-')
                    marks_lst.insert(2, '-')

            if marks_lst == []:
                pass
            else:
                l2.append(marks_lst)


            for p in y:
                if y[0].isdigit() and p.isalpha() and (p == 'PASS' or p == 'FAIL'):
                    pass_fail.append(p)

                if pass_fail == []:
                    pass
                else:
                    p_f.append(pass_fail)



            for tot in marks_lst:
                if tot!='-':
                    total += int(tot)
            total_lst.append(str(total))

            if total_lst != ['0']:
                totalmark.append(total_lst)

    for j in range(0, len(lst)):
        lk = lst[j] + l2[j] + totalmark[j] + p_f[j]
        final.append(lk)

    for s in range(len(final)):
        for total_lst in range(0, len(final) - s - 1):
            if final[total_lst][-2] < final[total_lst + 1][-2]:
                final[total_lst], final[total_lst + 1] = final[total_lst + 1], final[total_lst]

    if opt=='':

        eng_no, eng_avg, tamil_no, tamil_avg, hindi_no, hindi_avg, french_no, french_avg, math_no, math_avg, sci_no, sci_avg, soc_no, soc_avg = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        for x_1 in range(len(final)):
            if final[x_1][2] != '-':
                eng_avg += int(final[x_1][2])
                eng_no += 1
            if final[x_1][3] != '-':
                tamil_avg += int(final[x_1][3])
                tamil_no += 1
            if final[x_1][4] != '-':
                hindi_avg += int(final[x_1][4])
                hindi_no += 1
            if final[x_1][5] != '-':
                french_avg += int(final[x_1][5])
                french_no += 1
            if final[x_1][6] != '-':
                math_avg += int(final[x_1][6])
                math_no += 1
            if final[x_1][7] != '-':
                sci_avg += int(final[x_1][7])
                sci_no += 1
            if final[x_1][8] != '-':
                soc_avg += int(final[x_1][8])
                soc_no += 1

        avg_1 = round(eng_avg / eng_no, 2)
        if tamil_no != 0:
            avg_2 = round(tamil_avg / tamil_no, 2)
        else:
            avg_2 = 0
        if hindi_no != 0:
            avg_3 = round(hindi_avg / hindi_no, 2)
        else:
            avg_3 = 0
        if french_no != 0:
            avg_4 = round(french_avg / french_no, 2)
        else:
            avg_4 = 0
        avg_5 = round(math_avg / math_no, 2)
        avg_6 = round(sci_avg / sci_no, 2)
        avg_7 = round(soc_avg / soc_no, 2)

        avg.extend(['', 'Average:', avg_1, avg_2, avg_3, avg_4, avg_5, avg_6, avg_7])

        w = csv.writer(g, delimiter=',')
        w.writerow(c_name)
        w.writerows(final)
        w.writerow(avg)
        g.close()

    else:
        name_lst = []
        for i in x1:
            name_lst.append(i.strip())

        ds_lst=[]
        bd_lst=[]
        for i in final:
            if i[0] not in name_lst:
                ds_lst.append(i)
            elif i[0] in name_lst:
                bd_lst.append(i)

        def option(var,f_ext):
            eng_no, eng_avg, tamil_no, tamil_avg, hindi_no, hindi_avg, french_no, french_avg, math_no, math_avg, sci_no, sci_avg, soc_no, soc_avg = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
            for x_1 in range(len(var)):
                if var[x_1][2] != '-':
                    eng_avg += int(var[x_1][2])
                    eng_no += 1
                if var[x_1][3] != '-':
                    tamil_avg += int(var[x_1][3])
                    tamil_no += 1
                if var[x_1][4] != '-':
                    hindi_avg += int(var[x_1][4])
                    hindi_no += 1
                if var[x_1][5] != '-':
                    french_avg += int(var[x_1][5])
                    french_no += 1
                if var[x_1][6] != '-':
                    math_avg += int(var[x_1][6])
                    math_no += 1
                if var[x_1][7] != '-':
                    sci_avg += int(var[x_1][7])
                    sci_no += 1
                if var[x_1][8] != '-':
                    soc_avg += int(var[x_1][8])
                    soc_no += 1

            avg_1 = round(eng_avg / eng_no, 2)
            if tamil_no != 0:
                avg_2 = round(tamil_avg / tamil_no, 2)
            else:
                avg_2 = 0
            if hindi_no != 0:
                avg_3 = round(hindi_avg / hindi_no, 2)
            else:
                avg_3 = 0
            if french_no != 0:
                avg_4 = round(french_avg / french_no, 2)
            else:
                avg_4 = 0
            avg_5 = round(math_avg / math_no, 2)
            avg_6 = round(sci_avg / sci_no, 2)
            avg_7 = round(soc_avg / soc_no, 2)

            avg.extend(['', 'Average:', avg_1, avg_2, avg_3, avg_4, avg_5, avg_6, avg_7])

            w = csv.writer(f_ext, delimiter=',')
            w.writerow(c_name)
            w.writerows(var)
            w.writerow(avg)
            avg.clear()
            f_ext.close()

        option(ds_lst,ds);option(bd_lst,bd)


def convertG12(file,opt=''):
    global ds, bd, x1, g
    f = open(file)
    if opt=='':
        g = open(file[:-4]+'.csv', 'w+', newline='')
    else:
        ds = open(file[:-4] + '(DS).csv', 'w+', newline='')
        bd = open(file[:-4] + '(BD).csv', 'w+', newline='')
        lst = open(opt)
        x1 = lst.readlines()
    x = f.readlines()
    c_name = ['ROLL_NO', 'NAME', 'ENG', 'MATHS', 'PHY', 'CHE', 'C.S', 'APP.MATHS', 'BIO', 'ECO', 'ACC.', 'B.S', 'TOTAL', 'CUT-OFF', 'RESULT']
    lst = []
    l2 = []
    p_f = []
    totalmark = []
    cutoff = []
    avg = []
    final = []

    for i in range(len(x)):
        y = x[i].split()
        if y == []:
            pass
        else:
            roll = ''
            name = ''
            eng, maths, phy, che, cs, ap_mat, bio, eco, acc, bs = '-', '-', '-', '-', '-', '-', '-', '-', '-', '-'
            total = 0
            marks_lst = []
            pass_fail = []
            total_lst = []
            c_o=[]

            for j in y:
                name_rollno = [roll, name[1:]]
                if y[0].isdigit() and j.isalpha() and j != ('PASS' or 'FAIL'):
                    name += j + ' '
                elif y[0].isdigit() and j.isdigit() and len(j) > 4:
                    roll += j


            if len(name_rollno[0]) > 1 and len(name_rollno[1]) > 1:
                lst.append(name_rollno)

            if y[0].isdigit() and len(y[0]) > 4:
                sub_codes = []
                marks_lst = [eng,maths,phy,che,cs,ap_mat,bio,eco,acc,bs]

                for code in y:
                    if code.isdigit() and len(code) == 3:
                        sub_codes.append(code)

                y_1 = x[i+1].split()
                for pos in range(len(y_1)):
                    try:
                        if y_1[pos].isdigit()==False:
                            y_1.remove(y_1[pos])
                    except:
                        break

                for m in range(len(sub_codes)):
                    if sub_codes[m]=='301':
                        marks_lst[0]=y_1[m]
                    elif sub_codes[m]=='041':
                       marks_lst[1]=y_1[m]
                    elif sub_codes[m]=='042':
                        marks_lst[2]=y_1[m]
                    elif sub_codes[m]=='043':
                        marks_lst[3]=y_1[m]
                    elif sub_codes[m]=='083':
                        marks_lst[4]=y_1[m]
                    elif sub_codes[m]=='241':
                        marks_lst[5]=y_1[m]
                    elif sub_codes[m] == '044':
                        marks_lst[6]=y_1[m]
                    elif sub_codes[m] == '030':
                        marks_lst[7]=y_1[m]
                    elif sub_codes[m] == '055':
                        marks_lst[8]=y_1[m]
                    elif sub_codes[m] == '054':
                        marks_lst[9]=y_1[m]
                    else:
                        print('New sub_code found!',y[m])

            if marks_lst == []:
                pass
            else:
                l2.append(marks_lst)


            for p in y:
                if y[0].isdigit() and p.isalpha() and (p == 'PASS' or p == 'FAIL'):
                    pass_fail.append(p)

            if pass_fail == []:
                pass
            else:
                p_f.append(pass_fail)

            for tot in marks_lst:
                if tot!='-':
                    total += int(tot)
            total_lst.append(str(total))

            if total_lst != ['0']:
                totalmark.append(total_lst)

            if marks_lst == []:
                pass
            elif marks_lst[1]!='-' and marks_lst[2]!='-' and marks_lst[3]!='-':
                cut_off = int(marks_lst[1])+ round(int(marks_lst[2])/2) + round(int(marks_lst[3])/2)
                c_o.append(str(cut_off))
            else:
                c_o.append('-')

            if c_o == []:
                pass
            else:
               cutoff.append(c_o)

    for j in range(0, len(lst)):
        lk = lst[j] + l2[j] + totalmark[j] + cutoff[j] + p_f[j]
        final.append(lk)

    for s in range(len(final)):
        for total_lst in range(0, len(final) - s - 1):
            if final[total_lst][-3] < final[total_lst + 1][-3]:
                final[total_lst], final[total_lst + 1] = final[total_lst + 1], final[total_lst]

    if opt=='':

        eng_no, eng_avg, maths_no, maths_avg, phy_no, phy_avg, che_no, che_avg, cs_no, cs_avg, app_mat_no, app_mat_avg, bio_no, bio_avg, eco_no, eco_avg, acc_no, acc_avg, bs_no, bs_avg = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        for x_1 in range(len(final)):
            if final[x_1][2] != '-':
                eng_avg += int(final[x_1][2])
                eng_no += 1
            if final[x_1][3] != '-':
                maths_avg += int(final[x_1][3])
                maths_no += 1
            if final[x_1][4] != '-':
                phy_avg += int(final[x_1][4])
                phy_no += 1
            if final[x_1][5] != '-':
                che_avg += int(final[x_1][5])
                che_no += 1
            if final[x_1][6] != '-':
                cs_avg += int(final[x_1][6])
                cs_no += 1
            if final[x_1][7] != '-':
                app_mat_avg += int(final[x_1][7])
                app_mat_no += 1
            if final[x_1][8] != '-':
                bio_avg += int(final[x_1][8])
                bio_no += 1
            if final[x_1][9] != '-':
                eco_avg += int(final[x_1][9])
                eco_no += 1
            if final[x_1][10] != '-':
                acc_avg += int(final[x_1][10])
                acc_no += 1
            if final[x_1][11] != '-':
                bs_avg += int(final[x_1][11])
                bs_no += 1

        if eng_no != 0:
            avg_1 = round(eng_avg / eng_no, 2)
        else:
            avg_1 = 0
        if maths_no != 0:
            avg_2 = round(maths_avg / maths_no, 2)
        else:
            avg_2 = 0
        if phy_no != 0:
            avg_3 = round(phy_avg / phy_no, 2)
        else:
            avg_3 = 0
        if che_no != 0:
            avg_4 = round(che_avg / che_no, 2)
        else:
            avg_4 = 0
        if che_no != 0:
            avg_5 = round(cs_avg / cs_no, 2)
        else:
            avg_5 = 0
        if app_mat_no != 0:
            avg_6 = round(app_mat_avg / app_mat_no, 2)
        else:
            avg_6 = 0
        if bio_no != 0:
            avg_7 = round(bio_avg / bio_no, 2)
        else:
            avg_7 = 0
        if eco_no != 0:
            avg_8 = round(eco_avg / bio_no, 2)
        else:
            avg_8 = 0
        if acc_no != 0:
            avg_9 = round(acc_avg / bio_no, 2)
        else:
            avg_9 = 0
        if bs_no != 0:
            avg_10 = round(bs_avg / bio_no, 2)
        else:
            avg_10 = 0
        avg.extend(['', 'Average:', avg_1, avg_2, avg_3, avg_4, avg_5, avg_6, avg_7, avg_8, avg_9, avg_10])

        w = csv.writer(g, delimiter=',')
        w.writerow(c_name)
        w.writerows(final)
        w.writerow(avg)
        g.close()

    else:
        name_lst = []
        for i in x1:
            name_lst.append(i.strip())

        ds_lst = []
        bd_lst = []
        for i in final:
            if i[0] not in name_lst:
                ds_lst.append(i)

            elif i[0] in name_lst:
                bd_lst.append(i)

        def option(var, f_ext):
            eng_no, eng_avg, maths_no, maths_avg, phy_no, phy_avg, che_no, che_avg, cs_no, cs_avg, app_mat_no, app_mat_avg, bio_no, bio_avg, eco_no, eco_avg, acc_no, acc_avg, bs_no, bs_avg = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
            for x_1 in range(len(var)):
                if var[x_1][2] != '-':
                    eng_avg += int(var[x_1][2])
                    eng_no += 1
                if var[x_1][3] != '-':
                    maths_avg += int(var[x_1][3])
                    maths_no += 1
                if var[x_1][4] != '-':
                    phy_avg += int(var[x_1][4])
                    phy_no += 1
                if var[x_1][5] != '-':
                    che_avg += int(var[x_1][5])
                    che_no += 1
                if var[x_1][6] != '-':
                    cs_avg += int(var[x_1][6])
                    cs_no += 1
                if var[x_1][7] != '-':
                    app_mat_avg += int(var[x_1][7])
                    app_mat_no += 1
                if var[x_1][8] != '-':
                    bio_avg += int(var[x_1][8])
                    bio_no += 1
                if var[x_1][9] != '-':
                    eco_avg += int(var[x_1][9])
                    eco_no += 1
                if var[x_1][10] != '-':
                    acc_avg += int(var[x_1][10])
                    acc_no += 1
                if var[x_1][11] != '-':
                    bs_avg += int(var[x_1][11])
                    bs_no += 1

            if eng_no != 0:
                avg_1 = round(eng_avg / eng_no, 2)
            else:
                avg_1 = 0
            if maths_no != 0:
                avg_2 = round(maths_avg / maths_no, 2)
            else:
                avg_2 = 0
            if phy_no != 0:
                avg_3 = round(phy_avg / phy_no, 2)
            else:
                avg_3 = 0
            if che_no != 0:
                avg_4 = round(che_avg / che_no, 2)
            else:
                avg_4 = 0
            if che_no != 0:
                avg_5 = round(cs_avg / cs_no, 2)
            else:
                avg_5 = 0
            if app_mat_no != 0:
                avg_6 = round(app_mat_avg / app_mat_no, 2)
            else:
                avg_6 = 0
            if bio_no != 0:
                avg_7 = round(bio_avg / bio_no, 2)
            else:
                avg_7 = 0
            if eco_no != 0:
                avg_8 = round(eco_avg / bio_no, 2)
            else:
                avg_8 = 0
            if acc_no != 0:
                avg_9 = round(acc_avg / bio_no, 2)
            else:
                avg_9 = 0
            if bs_no != 0:
                avg_10 = round(bs_avg / bio_no, 2)
            else:
                avg_10 = 0
            avg.extend(['', 'Average:', avg_1, avg_2, avg_3, avg_4, avg_5, avg_6, avg_7, avg_8, avg_9, avg_10])

            w = csv.writer(f_ext, delimiter=',')
            w.writerow(c_name)
            w.writerows(var)
            w.writerow(avg)
            avg.clear()
            f_ext.close()

        option(ds_lst, ds);option(bd_lst, bd)