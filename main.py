import pandas as pd
import numpy as np
from tabulate import tabulate


class GeneralElective:
    def __init__(self, excelPath, timeTablePath):
        self.courses = []
        self.data = pd.read_excel(excelPath)
        self.headers = GeneralElective.getHeaders(excelPath)
        self.df = GeneralElective.frameData(self.data, self.headers)

        self.tdata = pd.read_excel(timeTablePath)
        self.theaders = GeneralElective.getHeaders(timeTablePath)
        self.tdf = GeneralElective.frameData(self.tdata, self.theaders)

    def displayTable(self):
        print(tabulate(self.df, headers=self.headers))

    def colorHtml(self):
        # initialize all time to zeros
        for weekday in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
            for row in range(13):
                self.tdf.at[row, weekday] = 0

        # increment one to each time segment for selected courses
        reference = {}
        for course in self.courses:
            index = np.where(self.df['No.'] == course)[0]
            timeLine = [str(i) for i in self.df['Time & Venue'].values[index]]
            for line in timeLine:
                day = line.split(" ")[0]
                time = line[line.index(" ") + 1:line.index("(")]
                timespan = self.expandTime(time)
                # print(day, time, timespan)
                for interval in timespan:
                    indexTime = np.where(self.tdf['Time'] == interval)[0][0]
                    self.tdf.at[indexTime, day] += 1
                    if (indexTime, day) in reference:
                        reference[(indexTime, day)].add(course)
                    else:
                        reference[(indexTime,day)] = {course}

        # display the reference for the courses on each cell
        print(reference)
        for key in reference:
            self.tdf.at[key[0],key[1]] = str(self.tdf.at[key[0],key[1]]) + str(reference[key])
            print(reference[key])

        # draw the timetable in html format (colorful)
        html = self.tdf.to_html()

        # color the cells in table according to their value
        next = 0
        while next <= html.rindex("<td>"):
            short = GeneralElective.extract(html, "<td>", "</td>", next)
            left = GeneralElective.getLeftIndex(html, "<td>", next)

            if short.split(".")[0].isdigit() and short.split(".")[1].split("{")[0].isdigit() \
                    and int(float(short.split(".")[0])) >= 1:
                density = int(float(short.split(".")[0]))
                red = str(10 + 50 * density) if 10 + 50 * density < 255 else str(255)
                green = str(255 - 50 * density) if 255 - 50 * density > 10 else str(10)
                blue = str(255 - 50 * density) if 255 - 50 * density > 10 else str(10)
                style = "<td style=\"background-color:rgba(" + red + ", " + green + "," + blue + ",0.5)\">"
                html = html[:left] + style + html[left + 4:]
                next = (left + 3 + len(short) + len(style))
            else:
                next = (left + 3 + len(short) + len("<td>"))

        # write the html to a file
        print(html)
        with open("table.html", "w") as htmlFile:
            htmlFile.write(html)
        print(tabulate(self.tdf, headers=self.theaders))

    def expandTime(self, time):
        timespan = time.split("-")
        last = timespan[-1]
        finish = int(timespan[-1][:timespan[-1].index(".")])
        start = int(timespan[0][:timespan[0].index(".")])
        startam = timespan[0][timespan[0].index(".00") + 3:]
        diff = finish - start if finish - start > 0 else finish + 12 - start

        newTime = start + 1 if start + 1 <= 12 else start + 1 - 12
        while diff >= 2:
            if 6 < newTime < 12 and startam == "am":
                timespan.append(str(newTime) + ".00am")
            else:
                timespan.append(str(newTime) + ".00pm")
            newTime = newTime + 1 if newTime + 1 <= 12 else newTime + 1 - 12
            diff -= 1

        timespan.remove(last)
        timespan.append(last)
        timespanlist = []
        for i in range(len(timespan) - 1):
            timespanlist.append(timespan[i] + "-" + timespan[i + 1])
        return timespanlist

    def addCourses(self, args):
        self.courses.clear()
        for i in range(len(args)):
            if 1 <= args[i] <= 96:
                self.courses.append(args[i])
            else:
                print("Invalid course indice:", i)

    def countCredit(self):
        total = 0
        for i in self.courses:
            index = np.where(self.df['No.'] == i)[0][0]
            value = self.df['Credit'].values[index]
            total += value
            print(value)
        print("The total Credit score is: " + str(total))

    @classmethod
    def science(cls):
        return cls("ge.xlsx", "timeTable.xlsx")

    @staticmethod
    def extract(string, left, right, start):
        return string[string.index(left, start) + len(left):string.index(right, start)]

    @staticmethod
    def getLeftIndex(string, left, start):
        return string.index(left, start)

    @staticmethod
    def frameData(data, headers):
        df = pd.DataFrame(data, columns=headers)
        df = df.fillna(method='ffill', axis=0)  # resolved updating missing row entities
        return df

    @staticmethod
    def getHeaders(excelPath):
        headers = str(pd.read_excel(excelPath).columns)
        headers = headers[headers.index("[") + 1: headers.rindex("]")]
        return [word.replace("\'", "").strip() for word in headers.split(",")]


# all = [i for i in range(1, 97)]

# ============================ Readme =======================================
# ||  This is an open source project to generate
# ||

cst = GeneralElective.science()
# cst.addCourses(all)
cst.addCourses([3,6,9])
cst.countCredit()
cst.colorHtml()
