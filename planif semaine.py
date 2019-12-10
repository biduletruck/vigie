import cProfile
import time as t
import pandas as pd
from datetime import datetime, timedelta, time, date
import re
import sys


myPlan = {"conseiller": "", "Date": ""}


def int_to_datetime(value):
    """
    Naively converts an int-based time representation to a datetime object.
    Today is used a the date part and is unsignificant but is required in
    order to perform arithmetics on times.
    """
    return datetime.combine(date.today(), time(hour=int(value / 100), minute=(value % 100)))


def parse_time_slice(value):
    match = re.match(
        r"^\D*((?:[0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9])\D+((?:[0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9])\D*$", value)
    if not match:
        raise Exception("Could not parse string")
    else:
        return (int(match.group(1).replace(":", "")),
                int(match.group(2).replace(":", "")))


def get_work_slices(amplitude, shift, breaks, interval, name, day):
    # Note: we don't care about the date here, but arithmetics on time objects can
    # only be achieved between full datetime and timedelta objects.

    start = int_to_datetime(amplitude[0])
    end = int_to_datetime(amplitude[1])
    shift_in = int_to_datetime(shift[0])
    shift_out = int_to_datetime(shift[1])
    delta = timedelta(minutes=interval)
    cbreaks = [(int_to_datetime(b[0]), int_to_datetime(b[1])) for b in breaks]
    time = start

    result_size = ((end - start) / delta)
    result = [False] * int(result_size)
    index = 0
    myPlan["conseiller"] = name
    myPlan["Date"] = day
    while time < end:
        working = int(0)
        if shift_in <= time < shift_out:
            working = int(1)
            for b in cbreaks:
                if b[0] <= time < b[1]:
                    working = int(0)

        result[index] = working

        # k = str(time.strftime("%H:%M"))
        # print("Time slice: " + str(time.strftime("%H:%M")) + " Working: " + str(result[index]))
        myPlan[time.strftime("%H:%M")] = result[index]

        # myPlan["conseiller"] = [name]
        # myPlan["Date"] = [day]
        # myPlan[time.strftime("%H:%M")] = [result[index]]

        time += delta
        index += 1

    return result


fileName = input("nom du fichier à analyser ? ... ")
planingName, ext = fileName.split(".")
excel = pd.read_excel(fileName)


print("Début du travail, merci de patienter ...")


start_time = t.time()
# try:
#     fileName
#     print(fileName)
#     planingName, ext = fileName.split(".")
#     excel = pd.read_excel(fileName)
# except FileNotFoundError:
#     exit("fichier non trouvé")
#     sys.exit()
# except ValueError:
#     exit("fichier non trouvé")
#     sys.exit()


# plannif = pd.read_excel(fileName)

sortie = {"nom": [], "date": [], "debutShift": [], "finShift": [], "debutPause1": [], "finPause1": [],
          "debutPause2": [], "finPause2": [], "debutPause3": [], "finPause3": [], "debutPause4": [],
          "finPause4": [], "debutPause5": [], "finPause5": [], "debutPause6": [], "finPause6": [],
          "debutPause7": [], "finPause7": [], "debutPause8": [], "finPause8": []}
agent = ''
debut = ''
fin = ""
planification = []

for i in range(1000000):
    if type(excel.values[i][0]) != float:
        if excel.values[i][0] == 'agent':
            df = pd.DataFrame(sortie)
            df.to_excel(str(planingName) + "_Global.xlsx", index=None)
            break
        elif (type(excel.values[i][0]) == str) & (type(excel.values[i][0]) != float):
            agent = excel.values[i][0]
        else:
            sortie["nom"].append(agent)
            sortie["date"].append(excel.values[i][0])
            if type(excel.values[i][2]) == str:
                try:
                    debutShift, finShift = excel.values[i][2].split(" - ")
                    debut = debutShift
                    fin = finShift
                except ValueError:
                    pass
            else:
                debut, fin = ("off", "off")

            sortie["debutShift"].append(debut)
            sortie["finShift"].append(fin)

            for val in range(3, 11, 1):
                if type(excel.values[i][val]) == str:
                    try:
                        debutShift, finShift = excel.values[i][val].split(" - ")
                        # debut = debutShift
                        if len(debutShift) >= 8:
                            debut = debutShift[4:]
                        else:
                            debut = debutShift
                        fin = finShift
                    except ValueError:
                        pass
                else:
                    debut, fin = ("", "")

                pause = val - 2
                sortie["debutPause" + str(pause)].append(debut)
                sortie["finPause" + str(pause)].append(fin)

# plannif = pd.read_excel(str(planingName) + "_Global.xlsx")
# plan = pd.array(planingName)
plannif = pd.read_excel(str(planingName) + "_Global.xlsx")

# print(plannif)
# plannif = pd.read_excel("['s44', 'xlsx']_Global.xlsx")
# print(plannif)
planning = {}
rangeAmplitude = {}
debutAmplitude = 800
finAmplitude = 2100
name = ''
day = ''

for i in range(10000000):
    try:
        name = str(plannif.values[i][0])
        day = plannif.values[i][1]
        debut = int(plannif.values[i][2].replace(':', ''))
        fin = int(plannif.values[i][3].replace(':', ''))
    except IndexError:
        break
    except TypeError:
        pass
    except ValueError:
        pass
    except AttributeError:
        pass

    breaks = []

    # print(plannif.values[i][0])
    if plannif.values[i][2] == "off":
        pass
    elif type(plannif.values[i][2]) == float:
        pass
    else:
        # print(str(plannif.values[i][1]) + " " + str(plannif.values[i][0]) + " " + plannif.values[i][2])
        name = str(plannif.values[i][0])
        day = plannif.values[i][1]
        debut = int(plannif.values[i][2].replace(':', ''))
        fin = int(plannif.values[i][3].replace(':', ''))

        for c in range(4, 11, 2):
            try:
                breaks.append(
                    tuple((int(plannif.values[i][c].replace(':', '')), int(plannif.values[i][c + 1].replace(':', '')))))
            except ValueError:
                pass
            except AttributeError:
                pass

        # print(breaks)
        get_work_slices(amplitude=(debutAmplitude, finAmplitude), shift=(debut, fin), breaks=breaks, interval=5,
                        name=name, day=day)

        if not "conseiller" in planning:
            for k in myPlan:
                planning[k] = []

            for k in planning:
                planning[k].append(myPlan[k])
        else:
            for k in planning:
                planning[k].append(myPlan[k])
            #
            # if not time.strftime("%H:%M") in myPlan:
            #
            #
            # else:
            #     for k in myPlan
            #         myPlan["conseiller"].append(name)
            #     myPlan["Date"].append(day)
            #     for k in
            #         myPlan[time.strftime("%H:%M")].append(result[ins50.xlsdex])

        # print(planning)

df2 = pd.DataFrame(planning)
df2.to_excel(str(planingName) + "_Détails.xlsx", index=None)

print("Travail terminé en  : %s secondes ---" % (t.time() - start_time))
# df2.to_excel(str("s44") + "_Détails.xlsx", index=None)
# breaks = [
#     (1000, 1030),
#     (1200, 1300),
#     (1600, 1630)
# ]
#
#
# def profile():
#     # Call our function 1000 times with an (extreme) interval of 1
#     for i in range(0, 54000):
#         get_work_slices(amplitude=(600, 2000), shift=(800, 1800), breaks=breaks, interval=1)
#
#
# cProfile.run('profile()')


