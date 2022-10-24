import docx
import re
import xlwings
import pandas as pd
from tabulate import tabulate

docRelation = {"HRD": ("HRS"), "HRS": ("PRS"), "PRS": ("URS", "RISK"), "HTR": ("HTP"), "HTP": ("HRD", "HRS"), \
               "SDS": ("BOLUS", "ACE", "AID"), "ACE": ("PRS", "TBV", "DER"), "BOULUS": ("PRS"), "AID": ("PRS", "DER"), \
               "SVAL": ("BOLUS", "ACE", "AID"), "SVATR": ("SVAL"), "UT": ("UNIT"),
               "INS": ("UNIT")}  # to be created by the GUI

docFile = {"HRD": "HDS_new_pump.docx",
           "HRS": "HRS_new_pump.docx",
           "HTP": "HTP_new_pump.docx",
           "HTR": "HTR_new_pump.docx",
           "PRS": "PRS_new_pump.docx",
           "RISK": "RiskAnalysis_Pump.docx",
           "SDS": "SDS_New_pump_x04.docx",
           "ACE": "SRS_ACE_Pump_X01.docx",
           "BOLUS": "SRS_BolusCalc_Pump_X04.docx",
           "SRS": "SRS_DosingAlgorithm_X03.docx",
           "SVAL": "SVaP_new_pump.docx",
           "SVATR": "SVaTR_new_pump.docx",
           "UT": "SVeTR_new_pump.docx", "URS": "URS_new_pump.docx"}


def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    fullText = [ele for ele in fullText if ele.strip()]
    return fullText


'''
def main():
hdsLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/HDS_new_pump.docx')
hrsLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/HRS_new_pump.docx')
htpLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/HTP_new_pump.docx')
htrLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/HTR_new_pump.docx')
prsLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/PRS_new_pump.docx')
riskLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/RiskAnalysis_Pump.docx')
sdsLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/SDS_New_pump_x04.docx')
srsAceLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/SRS_ACE_Pump_X01.docx')
srsBolusLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/SRS_BolusCalc_Pump_X04.docx')
srsDoseLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/SRS_DosingAlgorithm_X03.docx')
svapLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/SVaP_new_pump.docx')
svatrLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/SVaTR_new_pump.docx')
svetrLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/SVeTR_new_pump.docx')
ursLst = getText('C:/Users/steph/OneDrive/Desktop/Docs_Project/URS_new_pump.docx')
'''
hdsLst = getText('/Users/adrian/Desktop/SampledocsTandem/HDS_new_pump.docx')
hrsLst = getText('/Users/adrian/Desktop/SampledocsTandem/HRS_new_pump.docx')
htpLst = getText('/Users/adrian/Desktop/SampledocsTandem/HTP_new_pump.docx')
htrLst = getText('/Users/adrian/Desktop/SampledocsTandem/HTR_new_pump.docx')
prsLst = getText('/Users/adrian/Desktop/SampledocsTandem/PRS_new_pump.docx')
riskLst = getText('/Users/adrian/Desktop/SampledocsTandem/RiskAnalysis_Pump.docx')
sdsLst = getText('/Users/adrian/Desktop/SampledocsTandem/SDS_New_pump_x04.docx')
srsAceLst = getText('/Users/adrian/Desktop/SampledocsTandem/SRS_ACE_Pump_X01.docx')
srsBolusLst = getText('/Users/adrian/Desktop/SampledocsTandem/SRS_BolusCalc_Pump_X04.docx')
srsDoseLst = getText('/Users/adrian/Desktop/SampledocsTandem/SRS_DosingAlgorithm_X03.docx')
svapLst = getText('/Users/adrian/Desktop/SampledocsTandem/SVaP_new_pump.docx')
svatrLst = getText('/Users/adrian/Desktop/SampledocsTandem/SVaTR_new_pump.docx')
svetrLst = getText('/Users/adrian/Desktop/SampledocsTandem/SVeTR_new_pump.docx')
ursLst = getText('/Users/adrian/Desktop/SampledocsTandem/URS_new_pump.docx')

index = 0
ind = []
#index1 = 0
#ind1 = []

for t in hdsLst:
    if re.search('.*:HRD:', t):
        ind.append(index)
        tt = t
        y = re.findall('\S*:HRD:\S*', t)
        z = re.findall('\S*:HRS:\S*', t)
        tt = tt.replace(y[0], '')
        tt = tt.replace(z[0], '')
        tt = tt.strip()

        dict1 = {'Child Tag': [y],
                 'Info': [tt],
                 'Parent Tag': [z]
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in hrsLst:
    if re.search('.*:HRS:', u):
        ind.append(index)
        uu = u
        a = re.findall('\S*:HRS:\S*', u)
        b = re.findall('\S*:PRS:\S*', u)
        uu = uu.replace(a[0], '')
        uu = uu.replace(b[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a],
                 'Info': [uu],
                 'Parent Tag': [b]}
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

'''
for u in htpLst:
    if re.search('.*:HTP:', u):
        ind.append(index)
        uu = u
        a1 = re.findall('\S*:HTP:\S*', u)
        b1 = re.findall('\S*:HRS:\S*', u)
        uu = uu.replace(a1[0], '')
        uu = uu.replace(b1[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a1],
                'Info': [uu],
                'Parent Tag': [b1]}

        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

   elif re.search('.*:HTP:', u):
        ind.append(index)
        uu = u
        a1 = re.findall('\S*:HTP:\S*', u)
        b1 = re.findall('\S*:HRD:\S*', u)
        uu = uu.replace(a1[0], '')
        uu = uu.replace(b1[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a1],
                 'Info': [uu],
                 'Parent Tag': [b1]}

        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
'''

for u in htrLst:
    if re.search('.*:HTR:', u):
        ind.append(index)
        uu = u
        a3 = re.findall('\S*:HTR:\S*', u)
        b3 = re.findall('\S*:HTP:\S*', u)
        uu = uu.replace(a3[0], '')
        uu = uu.replace(b3[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a3],
                 'Info': [uu],
                 'Parent Tag': [b3]}
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)
'''
for u in prsLst:
    if re.search('.*:PRS:', u):
        ind.append(index)
        uu = u
        a4 = re.findall('\S*:PRS:\S*', u)
        b4 = re.findall('\S*:URS:\S*', u)
        # a5 = re.findall('\S*:PRS:\S*', u)
        # b5 = re.findall('\S*:RISK:\S*', u)
        uu = uu.replace(a4[0], '')
        uu = uu.replace(b4[0], '')
        # uu = uu.replace(a5[0], '')
        # uu = uu.replace(b5[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a4],
                 'Info': [uu],
                 'Parent Tag': [b4],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)
'''
for u in riskLst:
    if re.search('.*:RISK:', u):
        ind.append(index)
        uu = u
        a6 = re.findall('\S*:RISK:\S*', u)
        uu = uu.replace(a6[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a6],
                 'Info': [uu]
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in sdsLst:
    if re.search('.*:SDS:', u):
        ind.append(index)
        uu = u
        a7 = re.findall('\S*:SDS:\S*', u)
        b7 = re.findall('\S*:SRS:\S*', u)
        uu = uu.replace(a7[0], '')
        uu = uu.replace(b7[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a7],
                 'Info': [uu],
                 'Parent Tag': [b7],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in srsAceLst:
    if re.search('.*:SRS:', u):
        ind.append(index)
        uu = u
        a8 = re.findall('\S*:SRS:\S*', u)
        b8 = re.findall('\S*:PRS:\S*', u)
        uu = uu.replace(a8[0], '')
        uu = uu.replace(b8[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a8],
                 'Info': [uu],
                 'Parent Tag': [b8],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in srsBolusLst:
    if re.search('.*:SRS:', u):
        ind.append(index)
        uu = u
        a9 = re.findall('\S*:SRS:\S*', u)
        b9 = re.findall('\S*:PRS:\S*', u)
        uu = uu.replace(a9[0], '')
        uu = uu.replace(b9[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a9],
                 'Info': [uu],
                 'Parent Tag': [b9],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in srsDoseLst:
    if re.search('.*:SRS:', u):
        ind.append(index)
        uu = u
        a10 = re.findall('\S*:SRS:\S*', u)
        b10 = re.findall('\S*:DER:\S*', u)
        uu = uu.replace(a10[0], '')
        uu = uu.replace(b10[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a10],
                 'Info': [uu],
                 'Parent Tag': [b10],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in svapLst:
    if re.search('.*:SVAL:', u):
        ind.append(index)
        uu = u
        a11 = re.findall('\S*:SVAL:\S*', u)
        b11 = re.findall('\S*:SRS:\S*', u)
        uu = uu.replace(a11[0], '')
        uu = uu.replace(b11[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a11],
                 'Info': [uu],
                 'Parent Tag': [b11],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in svatrLst:
    if re.search('.*:SVATR:', u):
        ind.append(index)
        uu = u
        a12 = re.findall('\S*:SVATR:\S*', u)
        b12 = re.findall('\S*:SVAL:\S*', u)
        uu = uu.replace(a12[0], '')
        uu = uu.replace(b12[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a12],
                 'Info': [uu],
                 'Parent Tag': [b12],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in svetrLst:
    if re.search('.*:UT:', u):
        ind.append(index)
        uu = u
        a13 = re.findall('\S*:UT:\S*', u)
        b13 = re.findall('\S*:UNIT:\S*', u)
        uu = uu.replace(a13[0], '')
        uu = uu.replace(b13[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a13],
                 'Info': [uu],
                 'Parent Tag': [b13],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    elif re.search('.*:INS:', u):
        ind.append(index)
        uu = u
        a13 = re.findall('\S*:INS:\S*', u)
        b13 = re.findall('\S*:UNIT:\S*', u)
        uu = uu.replace(a13[0], '')
        uu = uu.replace(b13[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a13],
                 'Info': [uu],
                 'Parent Tag': [b13],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)

for u in ursLst:
    if re.search('.*:URS:', u):
        ind.append(index)
        uu = u
        a14 = re.findall('\S*:URS:\S*', u)
        uu = uu.replace(a14[0], '')
        uu = uu.strip()

        dict1 = {'Child Tag': [a14],
                 'Info': [uu],
                 }
        df = pd.DataFrame(dict1)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid'))
    index = index + 1
    print(ind)
# print(txtLst)
