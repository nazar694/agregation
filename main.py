import pandas as pd
import numpy as np
from pandas import io


def make_report_head(head):
    file_report.writelines('\n' + head + '\n')
    print('\n', head)


def make_report(txt, col_head, ind):
    file_report.write('\t' + txt + '\n')
    print('\t', txt)
    if col_head in marks_data:
        marks_data[col_head].update({ind: txt})
    else:
        marks_data[col_head] = {ind: txt}


def persentage(element, code=1):
    count = len(np.where(element == code)[0])
    return round(count / LEN * 100, 1)


def persentage_among_user(element, code=0):
    base = np.where(~np.isnan(element))[0]
    base_l = len(base)
    if code == 0:
        return round(base_l / LEN * 100, 1)

    count = len(np.where(element == code)[0])
    return round(count / base_l * 100, 1)


def persentage_among_two(element1, code1, element2, code2):
    base = np.where(element1 == code1)[0]
    base_l = len(base)
    finder = np.where(element2 == code2)[0]

    count = len(np.intersect1d(base, finder))
    return round(count * 100 / base_l, 1)


def numb(element):
    return len(np.where(~np.isnan(element))[0])


def raw_count(element, code=0):
    return len(np.where(element == code)[0])


data = pd.read_excel("DATA.xlsx")
LEN = len(data)
print(LEN)

global marks_data, file_report
marks_data = {}
file_report = open(r"REPORT.txt", "w+")


# AGE   Mean score
def age_mean(element):
    return sum(element) / LEN


make_report_head('AGE')
make_report(f"Mean score = {age_mean(data['Age'])}", 'AGE', 0)

# WARM-UP: STATEMENTS ABOUT TV ADS  Percentage
make_report_head('WARM-UP: STATEMENTS ABOUT TV ADS  Percentage')
make_report(f'There are too many ads on TV = {persentage(data["W1.tooManyAdverts"])}%',
            'WARM-UP: STATEMENTS ABOUT TV ADS', 0)
make_report(f'Some the ads are better than the TV programmes = {persentage(data["W1.betterThanTV"])}%',
            'WARM-UP: STATEMENTS ABOUT TV ADS', 1)
make_report(f'I’ve seen some really clever ads recently = {persentage(data["W1.cleverTVAdverts"])}%',
            'WARM-UP: STATEMENTS ABOUT TV ADS', 2)
make_report(f'Ads are often misleading = {persentage(data["W1.misleadingAdverts"])}%',
            'WARM-UP: STATEMENTS ABOUT TV ADS', 3)
make_report(f'There are some ads that I remember for a long time = {persentage(data["W1.longCherishAdverts"])}%',
            'WARM-UP: STATEMENTS ABOUT TV ADS', 4)

# ENJOYMENT	Percentage
make_report_head('ENJOYMENT	Percentage')
make_report(f'+5 - enjoy watching it a lot = {persentage(data["Enjoyment"], 6)}%', 'ENJOYMENT', 0)
make_report(f'+4 - quite enjoy/enjoy a little watching it = {persentage(data["Enjoyment"], 121)}%', 'ENJOYMENT', 1)
make_report(f'+3 - won\'t mind watching it = {persentage(data["Enjoyment"], 122)}%', 'ENJOYMENT', 2)
make_report(f'+2 - won\'t enjoy watching it much = {persentage(data["Enjoyment"], 123)}%', 'ENJOYMENT', 3)
make_report(f'+1 - won\'t enjoy watching it at all = {persentage(data["Enjoyment"], 124)}%', 'ENJOYMENT', 4)


# ENJOYMENT	Mean score
def enjoy_mean(element):
    mark_5 = len(np.where(element == 6)[0])
    mark_4 = len(np.where(element == 121)[0])
    mark_3 = len(np.where(element == 122)[0])
    mark_2 = len(np.where(element == 123)[0])
    mark_1 = len(np.where(element == 124)[0])
    return round((mark_5 * 5 + mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head('ENJOYMENT	Mean score')
make_report(f'Mean Score = {enjoy_mean(data["Enjoyment"])}', 'ENJOYMENT', 5)

# ACTIVE INVOLVEMENT Percentage
make_report_head('ACTIVE INVOLVEMENT Percentage')
make_report(f'Pleasant = {persentage(data["Activeinvolvementgroup.Pleasant"])}%', 'ACTIVE INVOLVEMENT', 0)
make_report(f'Interesting = {persentage(data["Activeinvolvementgroup.Interesting"])}%', 'ACTIVE INVOLVEMENT', 1)
make_report(f'Boring = {persentage(data["Activeinvolvementgroup.Boring"])}%', 'ACTIVE INVOLVEMENT', 2)
make_report(f'Irritating = {persentage(data["Activeinvolvementgroup.Irritating"])}%', 'ACTIVE INVOLVEMENT', 3)
make_report(f'Soothing = {persentage(data["Activeinvolvementgroup.Soothing"])}%', 'ACTIVE INVOLVEMENT', 4)
make_report(f'Distinctive = {persentage(data["Activeinvolvementgroup.Distinctive"])}%', 'ACTIVE INVOLVEMENT', 5)
make_report(f'Dull = {persentage(data["Activeinvolvementgroup.Dull"])}%', 'ACTIVE INVOLVEMENT', 6)
make_report(f'Unpleasant = {persentage(data["Activeinvolvementgroup.Unpleasant"])}%', 'ACTIVE INVOLVEMENT', 7)
make_report(f'Gentle = {persentage(data["Activeinvolvementgroup.Gentle"])}%', 'ACTIVE INVOLVEMENT', 8)
make_report(f'Involving = {persentage(data["Activeinvolvementgroup.Involving"])}%', 'ACTIVE INVOLVEMENT', 9)
make_report(f'Weak = {persentage(data["Activeinvolvementgroup.Weak"])}%', 'ACTIVE INVOLVEMENT', 10)
make_report(f'Disturbing = {persentage(data["Activeinvolvementgroup.Disturbing"])}%', 'ACTIVE INVOLVEMENT', 11)


def persentage_group(element, group_name):
    count = {}
    count['Pleasant'] = len(np.where(element['Activeinvolvementgroup.Pleasant'] == 1)[0])
    count['Soothing'] = len(np.where(element['Activeinvolvementgroup.Soothing'] == 1)[0])
    count['Gentle'] = len(np.where(element['Activeinvolvementgroup.Gentle'] == 1)[0])
    count['Interesting'] = len(np.where(element['Activeinvolvementgroup.Interesting'] == 1)[0])
    count['Distinctive'] = len(np.where(element['Activeinvolvementgroup.Distinctive'] == 1)[0])
    count['Involving'] = len(np.where(element['Activeinvolvementgroup.Involving'] == 1)[0])
    count['Boring'] = len(np.where(element['Activeinvolvementgroup.Boring'] == 1)[0])
    count['Dull'] = len(np.where(element['Activeinvolvementgroup.Dull'] == 1)[0])
    count['Weak'] = len(np.where(element['Activeinvolvementgroup.Weak'] == 1)[0])
    count['Irritating'] = len(np.where(element['Activeinvolvementgroup.Irritating'] == 1)[0])
    count['Unpleasant'] = len(np.where(element['Activeinvolvementgroup.Unpleasant'] == 1)[0])
    count['Disturbing'] = len(np.where(element['Activeinvolvementgroup.Disturbing'] == 1)[0])

    all_count = sum(count.values())

    if group_name == 'Passive Positive':
        group_count = count['Pleasant'] + count['Soothing'] + count['Gentle']

    if group_name == 'Active Positive':
        group_count = count['Interesting'] + count['Distinctive'] + count['Involving']

    if group_name == 'Passive Negative':
        group_count = count['Boring'] + count['Dull'] + count['Weak']

    if group_name == 'Active Negative':
        group_count = count['Irritating'] + count['Unpleasant'] + count['Disturbing']

    if group_name == 'Active':
        group_count = count['Interesting'] + count['Distinctive'] + count['Involving'] + count['Irritating'] + count['Unpleasant'] + count['Disturbing']

    if group_name == 'Passive':
        group_count = count['Pleasant'] + count['Soothing'] + count['Gentle'] + count['Boring'] + count['Dull'] + count['Weak']

    if group_name == 'Positive':
        group_count = count['Pleasant'] + count['Soothing'] + count['Gentle'] + count['Interesting'] + count['Distinctive'] + count['Involving']

    if group_name == 'Negative':
        group_count = count['Boring'] + count['Dull'] + count['Weak'] + count['Irritating'] + count['Unpleasant'] + count['Disturbing']

    if group_name == 'Mean Score':
        mark_4 = count['Pleasant'] + count['Soothing'] + count['Gentle']
        mark_3 = count['Interesting'] + count['Distinctive'] + count['Involving']
        mark_2 = count['Boring'] + count['Dull'] + count['Weak']
        mark_1 = count['Irritating'] + count['Unpleasant'] + count['Disturbing']

        return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / all_count, 1)

    return round(group_count * 100 / all_count, 1)


make_report(f'Passive Positive = {persentage_group(data, "Passive Positive")}%', 'ACTIVE INVOLVEMENT', 12)
make_report(f'Active Positive = {persentage_group(data, "Active Positive")}%', 'ACTIVE INVOLVEMENT', 13)
make_report(f'Passive Negative = {persentage_group(data, "Passive Negative")}%', 'ACTIVE INVOLVEMENT', 14)
make_report(f'Active Negative = {persentage_group(data, "Active Negative")}%', 'ACTIVE INVOLVEMENT', 15)
make_report(f'Active = {persentage_group(data, "Active")}%', 'ACTIVE INVOLVEMENT', 16)
make_report(f'Passive = {persentage_group(data, "Passive")}%', 'ACTIVE INVOLVEMENT', 17)
make_report(f'Positive = {persentage_group(data, "Positive")}%', 'ACTIVE INVOLVEMENT', 18)
make_report(f'Negative = {persentage_group(data, "Negative")}%', 'ACTIVE INVOLVEMENT', 19)


# ACTIVE INVOLVEMENT	Mean score

make_report_head('ACTIVE INVOLVEMENT	Mean score')
make_report(f'Mean score = {persentage_group(data, "Mean Score")}', 'ACTIVE INVOLVEMENT', 20)

# BRANDING	Percentage
make_report_head('BRANDING	Percentage')
make_report(
    f'+5 - couldn\'t fail to/help but remember the ad was for... = {persentage(data["Branding"], 7)}%',
    'BRANDING', 0)
make_report(
    f'+4 - quite/somewhat good at making you remember it is for...  = {persentage(data["Branding"], 162)}%',
    'BRANDING', 1)
make_report(
    f'+3 - not all that good at/just OK at making you remember it is for... = {persentage(data["Branding"], 163)}%',
    'BRANDING', 2)
make_report(f'+2 - could have been an ad for any brand of... = {persentage(data["Branding"], 164)}%',
            'BRANDING', 3)
make_report(f'+1 - could have been an ad for almost anything... = {persentage(data["Branding"], 165)}%',
            'BRANDING', 4)


# BRANDING	Mean score
def branding_mean(element):
    mark_5 = len(np.where(element == 7)[0])
    mark_4 = len(np.where(element == 162)[0])
    mark_3 = len(np.where(element == 163)[0])
    mark_2 = len(np.where(element == 164)[0])
    mark_1 = len(np.where(element == 165)[0])
    return round((mark_5 * 5 + mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head('BRANDING	Mean score')
make_report(f'Mean Score = {branding_mean(data["Branding"])}', 'BRANDING', 5)

# USERSHIP	Percentage
make_report_head('USERSHIP	Percentage')
make_report(f'+6 Use most often = {persentage(data["Fbrandusage"], 178)}%', 'USERSHIP', 0)
make_report(f'+5 Use it regularly = {persentage(data["Fbrandusage"], 179)}%', 'USERSHIP', 1)
make_report(f'+4 Use from time to time = {persentage(data["Fbrandusage"], 180)}%', 'USERSHIP', 2)
make_report(f'+3 Have tried in the past = {persentage(data["Fbrandusage"], 181)}%', 'USERSHIP', 3)
make_report(f'+2 Heard of but not tried = {persentage(data["Fbrandusage"], 182)}%', 'USERSHIP', 4)
make_report(f'+1 Never heard of until today = {persentage(data["Fbrandusage"], 183)}%', 'USERSHIP', 5)

# NUMBER OF USERS (PERCENTAGE) (ALL VARIANTS)	Percentage
make_report_head('NUMBER OF BRAND USAGE    Percentage')
make_report(f'Brand Usage = {persentage(data["Fbrandusage"], 178) + persentage(data["Fbrandusage"], 179)}%', 'USERSHIP',
            6)

# PERSUASION AMONG USERS	Percentage
make_report_head('PERSUASION AMONG USERS	Percentage')
make_report(f'+4 = {persentage_among_user(data["Fpersuasionuser"], 36)}%', 'PERSUASION AMONG USERS', 0)
make_report(f'+3 = {persentage_among_user(data["Fpersuasionuser"], 39)}%', 'PERSUASION AMONG USERS', 1)
make_report(f'+2 = {persentage_among_user(data["Fpersuasionuser"], 42)}%', 'PERSUASION AMONG USERS', 2)
make_report(f'+1 = {persentage_among_user(data["Fpersuasionuser"], 45)}%', 'PERSUASION AMONG USERS', 3)


# PERSUASION AMONG USERS	Mean score
def pau_mean(element):
    base = np.where(~np.isnan(element))[0]
    base_l = len(base)
    mark_4 = len(np.where(element == 36)[0])
    mark_3 = len(np.where(element == 39)[0])
    mark_2 = len(np.where(element == 42)[0])
    mark_1 = len(np.where(element == 45)[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION AMONG USERS	Mean score')
make_report(f'Mean Score = {pau_mean(data["Fpersuasionuser"])}', 'PERSUASION AMONG USERS', 4)

# NUMBER OF USERS (RAW COUNT)	Other
make_report_head('NUMBER OF USERS (RAW COUNT)	Other')
make_report(f'Other = {numb(data["Fpersuasionuser"])}', 'PERSUASION AMONG USERS', 5)

# NUMBER OF USERS (PERCENTAGE)	Percentage
make_report_head('NUMBER OF USERS (PERCENTAGE)	Percentage')
make_report(f'Users = {persentage_among_user(data["Fpersuasionuser"])}%', 'PERSUASION AMONG USERS', 6)

# PERSUASION: USE MOST OFTEN	Percentage
make_report_head('PERSUASION: USE MOST OFTEN	Percentage')
make_report(
    f'+4 - strongly encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 178, data["Fpersuasionuser"], 36)}%',
    'USE MOST OFTEN', 0)
make_report(
    f'+3 - encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 178, data["Fpersuasionuser"], 39)}%',
    'USE MOST OFTEN', 1)
make_report(
    f'+2 - makes no difference to whether I will go on using it = {persentage_among_two(data["Fbrandusage"], 178, data["Fpersuasionuser"], 42)}%',
    'USE MOST OFTEN', 2)
make_report(
    f'+1 - puts me off/makes me less likely to continue using it = {persentage_among_two(data["Fbrandusage"], 178, data["Fpersuasionuser"], 45)}%',
    'USE MOST OFTEN', 3)


# PERSUASION: USE MOST OFTEN	Mean score
def pumo_mean(base_element, element):
    base = np.where((base_element == 178))[0]
    base_l = len(base)
    mark_4 = len(np.where((base_element == 178) & (element == 36))[0])
    mark_3 = len(np.where((base_element == 178) & (element == 39))[0])
    mark_2 = len(np.where((base_element == 178) & (element == 42))[0])
    mark_1 = len(np.where((base_element == 178) & (element == 45))[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION: USE MOST OFTEN	Mean score')
make_report(f'Mean score = {pumo_mean(data["Fbrandusage"], data["Fpersuasionuser"])}', 'USE MOST OFTEN', 4)

# USE MOST OFTEN (RAW COUNT)	Other
make_report_head('USE MOST OFTEN (RAW COUNT)	Other')
make_report(f'Raw count  = {raw_count(data["Fbrandusage"], 178)}', 'USE MOST OFTEN', 5)

# USE MOST OFTEN (PERCENTAGE)	Percentage
make_report_head('USE MOST OFTEN (PERCENTAGE)	Percentage')
make_report(f'Sample percentage = {persentage(data["Fbrandusage"], 178)}%', 'USE MOST OFTEN', 6)

# PERSUASION: USE REGULARLY	Percentage
make_report_head('PERSUASION: USE REGULARLY	Percentage')
make_report(
    f'+4 - strongly encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 179, data["Fpersuasionuser"], 36)}%',
    'USE REGULARLY', 0)
make_report(
    f'+3 - encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 179, data["Fpersuasionuser"], 39)}%',
    'USE REGULARLY', 1)
make_report(
    f'+2 - makes no difference to whether I will go on using it = {persentage_among_two(data["Fbrandusage"], 179, data["Fpersuasionuser"], 42)}%',
    'USE REGULARLY', 2)
make_report(
    f'+1 - puts me off/makes me less likely to continue using it = {persentage_among_two(data["Fbrandusage"], 179, data["Fpersuasionuser"], 45)}%',
    'USE REGULARLY', 3)


# PERSUASION: USE REGULARLY	Mean score
def pur_mean(base_element, element):
    base = np.where((base_element == 179))[0]
    base_l = len(base)
    mark_4 = len(np.where((base_element == 179) & (element == 36))[0])
    mark_3 = len(np.where((base_element == 179) & (element == 39))[0])
    mark_2 = len(np.where((base_element == 179) & (element == 42))[0])
    mark_1 = len(np.where((base_element == 179) & (element == 45))[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION: USE REGULARLY	Mean score')
make_report(f'Mean score = {pur_mean(data["Fbrandusage"], data["Fpersuasionuser"])}', 'USE REGULARLY', 4)

# USE REGULARLY (RAW COUNT)	Other
make_report_head('USE REGULARLY (RAW COUNT)	Other')
make_report(f'Raw count  = {raw_count(data["Fbrandusage"], 179)}', 'USE REGULARLY', 5)

# USE REGULARLY (PERCENTAGE)	Percentage
make_report_head('USE REGULARLY (PERCENTAGE)	Percentage')
make_report(f'Sample percentage = {persentage(data["Fbrandusage"], 179)}%', 'USE REGULARLY', 6)

# PERSUASION AMONG TRIALLISTS	Percentage
make_report_head('PERSUASION AMONG TRIALLISTS	Percentage')
make_report(f'+4 = {persentage_among_user(data["Fpersuasiontriallist"], 38)}%', 'PERSUASION AMONG TRIALLISTS', 0)
make_report(f'+3 = {persentage_among_user(data["Fpersuasiontriallist"], 41)}%', 'PERSUASION AMONG TRIALLISTS', 1)
make_report(f'+2 = {persentage_among_user(data["Fpersuasiontriallist"], 45)}%', 'PERSUASION AMONG TRIALLISTS', 2)
make_report(f'+1 = {persentage_among_user(data["Fpersuasiontriallist"], 48)}%', 'PERSUASION AMONG TRIALLISTS', 3)


# PERSUASION AMONG TRIALLISTS	Mean score
def pat_mean(element):
    base = np.where(~np.isnan(element))[0]
    base_l = len(base)
    mark_4 = len(np.where(element == 38)[0])
    mark_3 = len(np.where(element == 41)[0])
    mark_2 = len(np.where(element == 45)[0])
    mark_1 = len(np.where(element == 48)[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION AMONG TRIALLISTS	Mean score')
make_report(f'Mean Score = {pat_mean(data["Fpersuasiontriallist"])}', 'PERSUASION AMONG TRIALLISTS', 4)

# NUMBER OF TRIALLISTS (RAW COUNT)	Other
make_report_head('NUMBER OF TRIALLISTS (RAW COUNT)	Other')
make_report(f'Other = {numb(data["Fpersuasiontriallist"])}', 'PERSUASION AMONG TRIALLISTS', 5)

# NUMBER OF TRIALLISTS (PERCENTAGE)	Percentage
make_report_head('NUMBER OF TRIALLISTS (PERCENTAGE)	Percentage')
make_report(f'Triallists = {persentage_among_user(data["Fpersuasiontriallist"])}%', 'PERSUASION AMONG TRIALLISTS', 6)

# PERSUASION: USE FROM TIME TO TIME	Percentage
make_report_head('PERSUASION: USE FROM TIME TO TIME	Percentage')
make_report(
    f'+4 - strongly encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 180, data["Fpersuasiontriallist"], 38)}%',
    'USE FROM TIME TO TIME', 0)
make_report(
    f'+3 - encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 180, data["Fpersuasiontriallist"], 41)}%',
    'USE FROM TIME TO TIME', 1)
make_report(
    f'+2 - makes no difference to whether I will go on using it = {persentage_among_two(data["Fbrandusage"], 180, data["Fpersuasiontriallist"], 45)}%',
    'USE FROM TIME TO TIME', 2)
make_report(
    f'+1 - puts me off/makes me less likely to continue using it = {persentage_among_two(data["Fbrandusage"], 180, data["Fpersuasiontriallist"], 48)}%',
    'USE FROM TIME TO TIME', 3)


# PERSUASION: USE FROM TIME TO TIME	Mean score
def pufttt_mean(base_element, element):
    base = np.where((base_element == 180))[0]
    base_l = len(base)
    mark_4 = len(np.where((base_element == 180) & (element == 38))[0])
    mark_3 = len(np.where((base_element == 180) & (element == 41))[0])
    mark_2 = len(np.where((base_element == 180) & (element == 45))[0])
    mark_1 = len(np.where((base_element == 180) & (element == 48))[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION: USE REGULARLY	Mean score')
make_report(f'Mean score = {pufttt_mean(data["Fbrandusage"], data["Fpersuasiontriallist"])}', 'USE FROM TIME TO TIME',
            4)

# USE FROM TIME TO TIME (RAW COUNT)	Other
make_report_head('USE FROM TIME TO TIME (RAW COUNT)	Other')
make_report(f'Raw count  = {raw_count(data["Fbrandusage"], 180)}', 'USE FROM TIME TO TIME', 5)

# USE FROM TIME TO TIME (PERCENTAGE)	Percentage
make_report_head('USE FROM TIME TO TIME (PERCENTAGE)	Percentage')
make_report(f'Sample percentage = {persentage(data["Fbrandusage"], 180)}%', 'USE FROM TIME TO TIME', 6)

# PERSUASION: TRIED IN THE PAST	Percentage
make_report_head('PERSUASION: TRIED IN THE PAST	Percentage')
make_report(
    f'+4 - strongly encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 181, data["Fpersuasiontriallist"], 38)}%',
    'TRIED IN THE PAST', 0)
make_report(
    f'+3 - encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 181, data["Fpersuasiontriallist"], 41)}%',
    'TRIED IN THE PAST', 1)
make_report(
    f'+2 - makes no difference to whether I will go on using it = {persentage_among_two(data["Fbrandusage"], 181, data["Fpersuasiontriallist"], 45)}%',
    'TRIED IN THE PAST', 2)
make_report(
    f'+1 - puts me off/makes me less likely to continue using it = {persentage_among_two(data["Fbrandusage"], 181, data["Fpersuasiontriallist"], 48)}%',
    'TRIED IN THE PAST', 3)


# PERSUASION: TRIED IN THE PAST	Mean score
def ptitp_mean(base_element, element):
    base = np.where((base_element == 181))[0]
    base_l = len(base)
    mark_4 = len(np.where((base_element == 181) & (element == 38))[0])
    mark_3 = len(np.where((base_element == 181) & (element == 41))[0])
    mark_2 = len(np.where((base_element == 181) & (element == 45))[0])
    mark_1 = len(np.where((base_element == 181) & (element == 48))[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION: TRIED IN THE PAST	Mean score')
make_report(f'Mean score = {ptitp_mean(data["Fbrandusage"], data["Fpersuasiontriallist"])}', 'TRIED IN THE PAST', 4)

# TRIED IN THE PAST (RAW COUNT)	Other
make_report_head('TRIED IN THE PAST (RAW COUNT)	Other')
make_report(f'Raw count  = {raw_count(data["Fbrandusage"], 181)}', 'TRIED IN THE PAST', 5)

# TRIED IN THE PAST (PERCENTAGE)	Percentage
make_report_head('TRIED IN THE PAST (PERCENTAGE)	Percentage')
make_report(f'Sample percentage = {persentage(data["Fbrandusage"], 181)}%', 'TRIED IN THE PAST', 6)

# PERSUASION AMONG NON USERS	Percentage
make_report_head('PERSUASION AMONG NON USERS	Percentage')
make_report(f'+4 = {persentage_among_user(data["Fpersuasionnontriallist"], 37)}%', 'PERSUASION AMONG NON USERS', 0)
make_report(f'+3 = {persentage_among_user(data["Fpersuasionnontriallist"], 40)}%', 'PERSUASION AMONG NON USERS', 1)
make_report(f'+2 = {persentage_among_user(data["Fpersuasionnontriallist"], 44)}%', 'PERSUASION AMONG NON USERS', 2)
make_report(f'+1 = {persentage_among_user(data["Fpersuasionnontriallist"], 47)}%', 'PERSUASION AMONG NON USERS', 3)


# PERSUASION AMONG NON USERS	Mean score
def panu_mean(element):
    base = np.where(~np.isnan(element))[0]
    base_l = len(base)
    mark_4 = len(np.where(element == 37)[0])
    mark_3 = len(np.where(element == 40)[0])
    mark_2 = len(np.where(element == 44)[0])
    mark_1 = len(np.where(element == 47)[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION AMONG NON USERS	Mean score')
make_report(f'Mean Score = {panu_mean(data["Fpersuasionnontriallist"])}', 'PERSUASION AMONG NON USERS', 4)

# NUMBER OF NON USERS (ACTUAL)	Other
make_report_head('NUMBER OF NON USERS (RAW COUNT)	Other')
make_report(f'Other = {numb(data["Fpersuasionnontriallist"])}', 'PERSUASION AMONG NON USERS', 5)

# NUMBER OF NON USERS (PERCENTAGE)	Percentage
make_report_head('NUMBER OF NON USERS (PERCENTAGE)	Percentage')
make_report(f'Triallists = {persentage_among_user(data["Fpersuasionnontriallist"])}%', 'PERSUASION AMONG NON USERS', 6)

# PERSUASION: HEARD OF BUT NEVER TRIED	Percentage
make_report_head('PERSUASION: HEARD OF BUT NEVER TRIED	Percentage')
make_report(
    f'+4 - strongly encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 182, data["Fpersuasionnontriallist"], 37)}%',
    'HEARD OF BUT NEVER TRIED', 0)
make_report(
    f'+3 - encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 182, data["Fpersuasionnontriallist"], 40)}%',
    'HEARD OF BUT NEVER TRIED', 1)
make_report(
    f'+2 - makes no difference to whether I will go on using it = {persentage_among_two(data["Fbrandusage"], 182, data["Fpersuasionnontriallist"], 44)}%',
    'HEARD OF BUT NEVER TRIED', 2)
make_report(
    f'+1 - puts me off/makes me less likely to continue using it = {persentage_among_two(data["Fbrandusage"], 182, data["Fpersuasionnontriallist"], 47)}%',
    'HEARD OF BUT NEVER TRIED', 3)


# PERSUASION: HEARD OF BUT NEVER TRIED	Mean score
def phobnt_mean(base_element, element):
    base = np.where((base_element == 182))[0]
    base_l = len(base)
    mark_4 = len(np.where((base_element == 182) & (element == 37))[0])
    mark_3 = len(np.where((base_element == 182) & (element == 40))[0])
    mark_2 = len(np.where((base_element == 182) & (element == 44))[0])
    mark_1 = len(np.where((base_element == 182) & (element == 47))[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION: HEARD OF BUT NEVER TRIED	Mean score')
make_report(f'Mean score = {phobnt_mean(data["Fbrandusage"], data["Fpersuasionnontriallist"])}',
            'HEARD OF BUT NEVER TRIED', 4)

# HEARD OF BUT NEVER TRIED (RAW COUNT)	Other
make_report_head('HEARD OF BUT NEVER TRIED (RAW COUNT)	Other')
make_report(f'Raw count  = {raw_count(data["Fbrandusage"], 182)}', 'HEARD OF BUT NEVER TRIED', 5)

# HEARD OF BUT NEVER TRIED (PERCENTAGE)	Percentage
make_report_head('HEARD OF BUT NEVER TRIED (PERCENTAGE)	Percentage')
make_report(f'Sample percentage = {persentage(data["Fbrandusage"], 182)}%', 'HEARD OF BUT NEVER TRIED', 6)

# PERSUASION: NEVER HEARD OF	Percentage
make_report_head('PERSUASION: NEVER HEARD OF	Percentage')
make_report(
    f'+4 - strongly encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 183, data["Fpersuasionnontriallist"], 37)}%',
    'NEVER HEARD OF', 0)
make_report(
    f'+3 - encourages me to go on using it = {persentage_among_two(data["Fbrandusage"], 183, data["Fpersuasionnontriallist"], 40)}%',
    'NEVER HEARD OF', 1)
make_report(
    f'+2 - makes no difference to whether I will go on using it = {persentage_among_two(data["Fbrandusage"], 183, data["Fpersuasionnontriallist"], 44)}%',
    'NEVER HEARD OF', 2)
make_report(
    f'+1 - puts me off/makes me less likely to continue using it = {persentage_among_two(data["Fbrandusage"], 183, data["Fpersuasionnontriallist"], 47)}%',
    'NEVER HEARD OF', 3)


# PERSUASION: NEVER HEARD OF	Mean score
def pnho_mean(base_element, element):
    base = np.where((base_element == 183))[0]
    base_l = len(base)
    mark_4 = len(np.where((base_element == 183) & (element == 37))[0])
    mark_3 = len(np.where((base_element == 183) & (element == 40))[0])
    mark_2 = len(np.where((base_element == 183) & (element == 44))[0])
    mark_1 = len(np.where((base_element == 183) & (element == 47))[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / base_l, 1)


make_report_head('PERSUASION: NEVER HEARD OF	Mean score')
make_report(f'Mean score = {pnho_mean(data["Fbrandusage"], data["Fpersuasionnontriallist"])}', 'NEVER HEARD OF', 4)

# NEVER HEARD OF (RAW COUNT)	Other
make_report_head('NEVER HEARD OF (RAW COUNT)	Other')
make_report(f'Raw count  = {raw_count(data["Fbrandusage"], 183)}', 'NEVER HEARD OF', 5)

# NEVER HEARD OF (PERCENTAGE)	Percentage
make_report_head('NEVER HEARD OF (PERCENTAGE)	Percentage')
make_report(f'Sample percentage = {persentage(data["Fbrandusage"], 183)}%', 'NEVER HEARD OF', 6)

# BRAND DIFFERENTIATION	Percentage
make_report_head('BRAND DIFFERENTIATION	Percentage')
make_report(f'+5 - Agree strongly = {persentage(data["Branddifference"], 424)}%', 'BRAND DIFFERENTIATION', 0)
make_report(f'+4 - Agree slightly = {persentage(data["Branddifference"], 425)}%', 'BRAND DIFFERENTIATION', 1)
make_report(f'+3 - Neither agree nor disagree = {persentage(data["Branddifference"], 426)}%', 'BRAND DIFFERENTIATION',
            2)
make_report(f'+2 - Disagree slightly = {persentage(data["Branddifference"], 427)}%', 'BRAND DIFFERENTIATION', 3)
make_report(f'+1 - Disagree strongly = {persentage(data["Branddifference"], 428)}%', 'BRAND DIFFERENTIATION', 4)


# BRAND DIFFERENTIATION	Mean score
def brandinfferentiation_mean(element):
    mark_5 = len(np.where(element == 424)[0])
    mark_4 = len(np.where(element == 425)[0])
    mark_3 = len(np.where(element == 426)[0])
    mark_2 = len(np.where(element == 427)[0])
    mark_1 = len(np.where(element == 428)[0])
    return round((mark_5 * 5 + mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head('BRAND DIFFERENTIATION	Mean score')
make_report(f'Mean Score = {brandinfferentiation_mean(data["Branddifference"])}', 'BRAND DIFFERENTIATION', 5)


# Affinity	Percentage
def affinity_persentage(element, min_grade, max_grade):
    count = 0
    for el in element:
        if min_grade <= el <= max_grade:
            count += 1

    return round(count * 100 / LEN, 1)


make_report_head('Affinity	Percentage')
make_report(f'91-100 = {affinity_persentage(data["Affinity._1.slice"], 91, 100)}%', 'AFFINITY', 0)
make_report(f'81-90 = {affinity_persentage(data["Affinity._1.slice"], 81, 90)}%', 'AFFINITY', 1)
make_report(f'71-80 = {affinity_persentage(data["Affinity._1.slice"], 71, 80)}%', 'AFFINITY', 2)
make_report(f'61-70 = {affinity_persentage(data["Affinity._1.slice"], 61, 70)}%', 'AFFINITY', 3)
make_report(f'51-60 = {affinity_persentage(data["Affinity._1.slice"], 51, 60)}%', 'AFFINITY', 4)
make_report(f'41-50 = {affinity_persentage(data["Affinity._1.slice"], 41, 50)}%', 'AFFINITY', 5)
make_report(f'31-40 = {affinity_persentage(data["Affinity._1.slice"], 31, 40)}%', 'AFFINITY', 6)
make_report(f'21-30 = {affinity_persentage(data["Affinity._1.slice"], 21, 30)}%', 'AFFINITY', 7)
make_report(f'11-20 = {affinity_persentage(data["Affinity._1.slice"], 11, 20)}%', 'AFFINITY', 8)
make_report(f'1-10 = {affinity_persentage(data["Affinity._1.slice"], 1, 10)}%', 'AFFINITY', 9)


# Affinity	Mean score
def affinity_mean(element):
    summa = sum(element)
    return round(summa / LEN, 1)


make_report_head('Affinity	Mean score')

make_report(f'Mean score = {affinity_mean(data["Affinity._1.slice"])}', 'AFFINITY', 10)

# MESSAGE CHECK - 1	Percentage
make_report_head('MESSAGE CHECK - 1	Percentage')
make_report(f'Very likely to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_1.slice"], 8532)}%',
            'MESSAGE CHECK - 1', 0)
make_report(f'Quite likely to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_1.slice"], 8533)}%',
            'MESSAGE CHECK - 1', 1)
make_report(
    f'Probably wouldn\'t to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_1.slice"], 8534)}%',
    'MESSAGE CHECK - 1', 2)
make_report(
    f'Definitely wouldn\'t to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_1.slice"], 8535)}%',
    'MESSAGE CHECK - 1', 3)


# MESSAGE CHECK - 1	Mean score
def mess1_mean(element):
    mark_4 = len(np.where(element == 8532)[0])
    mark_3 = len(np.where(element == 8533)[0])
    mark_2 = len(np.where(element == 8534)[0])
    mark_1 = len(np.where(element == 8535)[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head('MESSAGE CHECK - 1	Mean score')
make_report(f'Mean Score = {mess1_mean(data["Messagecheck.messagecheckdriver_1.slice"])}', 'MESSAGE CHECK - 1', 4)

# MESSAGE CHECK - 2	Percentage
make_report_head('MESSAGE CHECK - 2	Percentage')
make_report(f'Very likely to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_2.slice"], 8537)}%',
            'MESSAGE CHECK - 2', 0)
make_report(f'Quite likely to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_2.slice"], 8538)}%',
            'MESSAGE CHECK - 2', 1)
make_report(
    f'Probably wouldn\'t to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_2.slice"], 8539)}%',
    'MESSAGE CHECK - 2', 2)
make_report(
    f'Definitely wouldn\'t to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_2.slice"], 8540)}%',
    'MESSAGE CHECK - 2', 3)


# MESSAGE CHECK - 2	Mean score
def mess2_mean(element):
    mark_4 = len(np.where(element == 8537)[0])
    mark_3 = len(np.where(element == 8538)[0])
    mark_2 = len(np.where(element == 8539)[0])
    mark_1 = len(np.where(element == 8540)[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head('MESSAGE CHECK - 2	Mean score')
make_report(f'Mean Score = {mess2_mean(data["Messagecheck.messagecheckdriver_2.slice"])}', 'MESSAGE CHECK - 2', 4)

# MESSAGE CHECK - 3	Percentage
make_report_head('MESSAGE CHECK - 3	Percentage')
make_report(f'Very likely to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_3.slice"], 8542)}%',
            'MESSAGE CHECK - 3', 0)
make_report(f'Quite likely to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_3.slice"], 8543)}%',
            'MESSAGE CHECK - 3', 1)
make_report(
    f'Probably wouldn\'t to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_3.slice"], 8544)}%',
    'MESSAGE CHECK - 3', 2)
make_report(
    f'Definitely wouldn\'t to stick in mind = {persentage(data["Messagecheck.messagecheckdriver_3.slice"], 8545)}%',
    'MESSAGE CHECK - 3', 3)


# MESSAGE CHECK - 3	Mean score
def mess3_mean(element):
    mark_4 = len(np.where(element == 8542)[0])
    mark_3 = len(np.where(element == 8543)[0])
    mark_2 = len(np.where(element == 8544)[0])
    mark_1 = len(np.where(element == 8545)[0])
    return round((mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head('MESSAGE CHECK - 3	Mean score')
make_report(f'Mean Score = {mess3_mean(data["Messagecheck.messagecheckdriver_3.slice"])}', 'MESSAGE CHECK - 3', 4)

# GENDER EQUALITY MEASUREMENT - FEMALE - Positive image of the female character/s that sets a good example for others	Percentage
make_report_head(
    'GENDER EQUALITY MEASUREMENT - FEMALE - Positive image of the female character/s that sets a good example for others	Percentage')
make_report(
    f'Positive = {persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11788) + persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11789)}%',
    'GENDER EQUALITY MEASUREMENT - FEMALE', 0)
make_report(f'+5 = {persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11788)}%',
            'GENDER EQUALITY MEASUREMENT - FEMALE', 1)
make_report(f'+4 = {persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11789)}%',
            'GENDER EQUALITY MEASUREMENT - FEMALE', 2)
make_report(f'+3 = {persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11790)}%',
            'GENDER EQUALITY MEASUREMENT - FEMALE', 3)
make_report(f'+2 = {persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11791)}%',
            'GENDER EQUALITY MEASUREMENT - FEMALE', 4)
make_report(f'+1 = {persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11792)}%',
            'GENDER EQUALITY MEASUREMENT - FEMALE', 5)
make_report(
    f'Negative = {persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11791) + persentage(data["Genderportrayal1.PositiveImageFemale.slice"], 11792)}%',
    'GENDER EQUALITY MEASUREMENT - FEMALE', 6)


# GENDER EQUALITY MEASUREMENT - FEMALE - Positive image of the female character/s that sets a good example for others	Mean score
def genderfemale_mean(element):
    mark_5 = len(np.where(element == 11788)[0])
    mark_4 = len(np.where(element == 11789)[0])
    mark_3 = len(np.where(element == 11790)[0])
    mark_2 = len(np.where(element == 11791)[0])
    mark_1 = len(np.where(element == 11792)[0])
    return round((mark_5 * 5 + mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head(
    'GENDER EQUALITY MEASUREMENT - FEMALE - Positive image of the female character/s that sets a good example for others	Mean score')
make_report(f'Mean Score = {genderfemale_mean(data["Genderportrayal1.PositiveImageFemale.slice"])}',
            'GENDER EQUALITY MEASUREMENT - FEMALE', 7)

# GENDER EQUALITY MEASUREMENT - MALE - Positive image of the male character/s that sets a good example for others	Percentage
make_report_head(
    'GENDER EQUALITY MEASUREMENT - MALE - Positive image of the female character/s that sets a good example for others	Percentage')
make_report(
    f'Positive = {persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11794) + persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11795)}%',
    'GENDER EQUALITY MEASUREMENT - MALE', 0)
make_report(f'+5 = {persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11794)}%',
            'GENDER EQUALITY MEASUREMENT - MALE', 1)
make_report(f'+4 = {persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11795)}%',
            'GENDER EQUALITY MEASUREMENT - MALE', 2)
make_report(f'+3 = {persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11796)}%',
            'GENDER EQUALITY MEASUREMENT - MALE', 3)
make_report(f'+2 = {persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11797)}%',
            'GENDER EQUALITY MEASUREMENT - MALE', 4)
make_report(f'+1 = {persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11798)}%',
            'GENDER EQUALITY MEASUREMENT - MALE', 5)
make_report(
    f'Negative = {persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11797) + persentage(data["Genderportrayal2.PositiveImageMale.slice"], 11798)}%',
    'GENDER EQUALITY MEASUREMENT - MALE', 6)


# GENDER EQUALITY MEASUREMENT - MALE - Positive image of the male character/s that sets a good example for others	Mean score
def gendermale_mean(element):
    mark_5 = len(np.where(element == 11794)[0])
    mark_4 = len(np.where(element == 11795)[0])
    mark_3 = len(np.where(element == 11796)[0])
    mark_2 = len(np.where(element == 11797)[0])
    mark_1 = len(np.where(element == 11798)[0])
    return round((mark_5 * 5 + mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head(
    'GENDER EQUALITY MEASUREMENT - MALE - Positive image of the female character/s that sets a good example for others	Mean score')
make_report(f'Mean Score = {gendermale_mean(data["Genderportrayal2.PositiveImageMale.slice"])}',
            'GENDER EQUALITY MEASUREMENT - MALE', 7)

# INCLUSION - The way people are presented represents a modern and progressive view of society	Percentage
make_report_head(
    'INCLUSION - The way people are presented represents a modern and progressive view of society	Percentage')
make_report(f'+5 Strongly agree = {persentage(data["Inclusion._1.slice"], 12696)}%', 'INCLUSION', 0)
make_report(f'+4 Somewhat agree = {persentage(data["Inclusion._1.slice"], 12697)}%', 'INCLUSION', 1)
make_report(f'+3 Neither agree nor disagree = {persentage(data["Inclusion._1.slice"], 12698)}%', 'INCLUSION', 2)
make_report(f'+2 Somewhat disagree = {persentage(data["Inclusion._1.slice"], 12699)}%', 'INCLUSION', 3)
make_report(f'+1 Strongly disagree = {persentage(data["Inclusion._1.slice"], 12700)}%', 'INCLUSION', 4)


# INCLUSION - The way people are presented represents a modern and progressive view of society	Mean score
def inclusion_mean(element):
    mark_5 = len(np.where(element == 12696)[0])
    mark_4 = len(np.where(element == 12697)[0])
    mark_3 = len(np.where(element == 12698)[0])
    mark_2 = len(np.where(element == 12699)[0])
    mark_1 = len(np.where(element == 12700)[0])
    return round((mark_5 * 5 + mark_4 * 4 + mark_3 * 3 + mark_2 * 2 + mark_1 * 1) / LEN, 1)


make_report_head(
    'INCLUSION - The way people are presented represents a modern and progressive view of society	Mean score')
make_report(f'Mean Score = {inclusion_mean(data["Inclusion._1.slice"])}', 'INCLUSION', 5)

# DEVICE TYPE USED	Percentage
make_report_head('DEVICE TYPE USED	Percentage')
make_report(f'PC/laptop/netbook = {persentage(data["DeviceDetails.kantarDevice"], 7583)}%', 'DEVICE TYPE USED', 0)
make_report(f'Large tablet = {persentage(data["DeviceDetails.kantarDevice"], 7584)}%', 'DEVICE TYPE USED', 1)
make_report(f'Medium tablet = {persentage(data["DeviceDetails.kantarDevice"], 7585)}%', 'DEVICE TYPE USED', 2)
make_report(f'Small tablet = {persentage(data["DeviceDetails.kantarDevice"], 7586)}%', 'DEVICE TYPE USED', 3)
make_report(f'Smart phone touch = {persentage(data["DeviceDetails.kantarDevice"], 7587)}%', 'DEVICE TYPE USED', 4)
make_report(f'Large Feature phone = {persentage(data["DeviceDetails.kantarDevice"], 7588)}%', 'DEVICE TYPE USED', 5)
make_report(f'Small Feature phone = {persentage(data["DeviceDetails.kantarDevice"], 7589)}%', 'DEVICE TYPE USED', 6)
make_report(f'Other = {persentage(data["DeviceDetails.kantarDevice"], 7590)}%', 'DEVICE TYPE USED', 7)

exel_marks_data = pd.DataFrame(marks_data)

file_name = 'REPORT.xlsx'
exel_marks_data.to_excel(file_name)

file_report.close()
print(data.info())
