#!/usr/bin/env python2.7

"""Parses responses from Survey Monkey survey and creates preliminary groups"""

import operator as op
import random
import xlrd

__author__ = 'Palmer Paul'
__version__ = '1.0.2'
__email__ = 'pzpaul2002@yahoo.com'

FILENAME = 'Passover Workshop Test Sheet (Palmer).xls'  # path to file
MAX_PER_GROUP = 10  # maximum number of people in a group
NUM_RANKED = 5  # i.e. rank the choices from one to...


class ExcelParser(object):
    """Loads and extracts data from Excel file. Also parses certain data."""
    def __init__(self, fname):
        self.fname = fname
        self.book = None
        self.data = None

    def load_data(self):
        """Load data from first sheet"""
        with xlrd.open_workbook(self.fname) as book:
            self.book = book
            self.data = book.sheet_by_index(0)

    def rows(self):
        """Make an iterable for the rows of a sheet"""
        if self.data is None:
            self.load_data()
        for i in xrange(self.data.nrows):
            yield self.data.row(i)

    def __extract_ranking_from_cell(self, cell):
        """Given a non-empty cell `cell`, return an integer of the ranking"""
        val = cell.value
        i = 0
        while i < len(val):
            if not val[i].isdigit():
                break
            i += 1
        return int(val[:i]) - 1

    def parse_row_to_student(self, row):
        """Return a new Student() based on the information in `row`"""
        iter_row = iter(row)

        first_name = next(iter_row).value
        last_name = next(iter_row).value
        grade = int(next(iter_row).value)

        choices = {}
        for i, cell in enumerate(iter_row):
            if cell.ctype != 0:
                ranking = self.__extract_ranking_from_cell(cell)
                choices[ranking] = i
        choices_arr = map(op.itemgetter(1), sorted(choices.iteritems()))

        return Student(first_name, last_name, grade, choices_arr)

    def get_choice_names(self):
        """returns ordered list of names of choices"""
        if self.data is None:
            self.load_data()
        return map(op.attrgetter('value'), self.data.row(1)[3:])


class Student(object):
    """Holds information about an individual student.
    Says how to display student information"""
    def __init__(self, first, last, grade, choices):
        self.first = first
        self.last = last
        self.grade = grade
        self.choices = choices
        self.group = None

    def __repr__(self):
        return 'Student({}, {}, {}, {})'.format(
            self.first, self.last, self.grade, self.choices)

    def __str__(self):
        return '{} {}'.format(
            self.first, self.last)


class Students(object):
    """Collection of Student objects.
    Provides methods to modify and display the collection"""
    def __init__(self, students=None):
        self.students = students if students is not None else []

    def __repr__(self):
        return 'Students({})'.format(self.students)

    def __str__(self):
        self.students.sort()
        return '\n'.join(str(student) for student in self.students)

    def __iter__(self):
        return iter(self.students)

    def add_student(self, student):
        """Add a student to the collection"""
        self.students.append(student)

    def size(self):
        """Return the number of students in the collection"""
        return len(self.students)

    def sort(self):
        """Sort the students by grade, then last name, then first name"""
        self.students = sorted(self.students, key=op.attrgetter(
            'grade', 'last', 'first'))

    def randomize(self):
        """Randomizes the order of the students in the collection"""
        random.shuffle(self.students)


class Application(object):
    """Main "driver" class. Makes everything work together."""
    def __init__(self, filename, max_per_group, num_ranked):
        excel_data = ExcelParser(filename)
        excel_data.load_data()

        self.choice_names = excel_data.get_choice_names()

        rows = excel_data.rows()
        next(rows)
        next(rows)
        self.all_students = Students(
            map(excel_data.parse_row_to_student, rows))

        self.make_groups(max_per_group, num_ranked)

    def make_groups(self, max_per_group, num_ranked):
        """
        Make groups based on students' preferences.
        Goes through all of the students and gives them
        their first choice if there is space in that group.
        Afterwards, goes through remaining students
        and tries to give them their second choice.
        This process continues until all students are assigned a group.
        """
        group_counts = [0] * len(self.choice_names)
        self.all_students.randomize()
        for n in xrange(num_ranked):
            for student in self.all_students:
                if student.group is None:
                    try:
                        choice = student.choices[n]
                    except IndexError:
                        pass
                    else:
                        if group_counts[choice] < max_per_group:
                            student.group = choice
                            group_counts[choice] += 1

    def query_students(self, query):
        """Find students who match a certain criterion"""
        return Students(filter(query, self.all_students))

    def print_groups(self):
        """Output the group assignments by group"""
        for i, group_name in enumerate(self.choice_names):
            print group_name.upper()
            students = self.query_students(lambda s: s.group == i)
            students.sort()
            for student in students:
                print '{} ({}th)'.format(student, student.grade)
            print

    def print_grades(self):
        """Output the group assignments by grade"""
        grades = (('Freshmen', 9), ('Sophomores', 10),
                  ('Juniors', 11), ('Seniors', 12))

        for grade in grades:
            print grade[0].upper()
            students = self.query_students(lambda s: s.grade == grade[1])
            students.sort()
            for student in students:
                print str(student) + ': ' + (self.choice_names[student.group]
                                             if student.group is not None
                                             else 'None')
            print

    def main(self):
        """It all starts here! Entrypoint for program."""
        self.print_groups()
        self.print_grades()

if __name__ == '__main__':
    app = Application(FILENAME, MAX_PER_GROUP, NUM_RANKED)
    app.main()
