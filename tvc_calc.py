from datetime import datetime

from logging_tvc import logger
from tvc_edit import today as today


def change_to_hours(number: float):
    """Change float number into hours and minutes format
        number: float value from a report"""
    h = int(number)
    m: int = int(round((number - h) * 60))
    return h, m

def tvc_calc(date_week: datetime, date_calc = datetime.today()):
    """Calculates TVC value that should be claimed

    Keyword arguments:
         date_week -- date of end (friday) of the week for calculations
         date_calc -- date for which we want to check tvc data, default: day of running the script"""

    monday = date_week[0]
    if date_calc > date_week[1]:  # friday
        # date_calc is after friday -> tvc = 40
        tvc = 40
    elif date_calc >= monday <= date_week[1]:
        # date_calc is in the current week -> tvc < 40
        weekday = datetime.isoweekday(date_calc)
        tvc = (weekday - 1) * 8
    else:
        # date_calc is before friday - future case
        tvc = 0
    return tvc

def tvc_miss(data, weeks):
    """Calculates how many hours (and minutes) are missing for each week"""

    tvc_report_data = []

    for w in weeks:
        tvc_val = tvc_calc(w, today)
        logger.info(f'Proper value to claim: {tvc_val} - week {" - ".join(map(str,w))}')
        if tvc_val == 0:
            logger.warn(f'TVC not calculated for week {" - ".join(map(str,w))} -> in the future')

        for entry in data:
            tvc_claim = entry[str(w[1])]
            hour, _ = change_to_hours(float(tvc_claim))

            if tvc_val > 0 and hour < (tvc_val + 3):
                hour_miss, min_miss = change_to_hours(float(tvc_val) - tvc_claim)
                if hour_miss > 1:
                    tvc_report_data.append({
                        'email': entry['email_id'],
                        'username': entry['username'].split('/')[0].title(),
                        'week': w[1],
                        'tvc_val': tvc_val,
                        'tvc_claim': tvc_claim,
                        'hour_miss': hour_miss,
                        'min_miss': min_miss
                    })
    return tvc_report_data