import mysql.connector
import xlsxwriter

class Exporter:
    def write_subvoyage(self, id):
        self.row += 1
        self.worksheet.write(self.row, self.col+1, "SUBVOYAGES:", self.bold)
        self.row+= 1
        sv_col = self.col + 1

        self.worksheet.write(self.row, sv_col, 'subvoyage_id', self.bold)
        self.worksheet.write(self.row, sv_col + 1, 'subvoyage_type', self.bold)
        self.worksheet.write(self.row, sv_col + 2, 'sub_dept_location', self.bold)
        self.worksheet.write(self.row, sv_col + 3, 'sub_dept_location_standardized', self.bold)
        self.worksheet.write(self.row, sv_col + 4, 'sub_dept_location_status', self.bold)
        self.worksheet.write(self.row, sv_col + 5, 'sub_dept_date_as_source', self.bold)
        self.worksheet.write(self.row, sv_col + 6, 'sub_dept_date_as_source', self.bold)
        self.worksheet.write(self.row, sv_col + 7, 'sub_dept_date_year', self.bold)
        self.worksheet.write(self.row, sv_col + 8, 'sub_dept_date_month', self.bold)
        self.worksheet.write(self.row, sv_col + 9, 'sub_dept_date_day', self.bold)
        self.worksheet.write(self.row, sv_col + 10, 'sub_dept_date_year_to', self.bold)
        self.worksheet.write(self.row, sv_col + 11, 'sub_dept_date_month_to', self.bold)
        self.worksheet.write(self.row, sv_col + 12, 'sub_dept_date_day_to', self.bold)
        self.worksheet.write(self.row, sv_col + 13, 'sub_dept_date_status', self.bold)
        self.worksheet.write(self.row, sv_col + 14, 'sub_dept_date_relative', self.bold)
        self.worksheet.write(self.row, sv_col + 15, 'sub_arrival_location', self.bold)
        self.worksheet.write(self.row, sv_col + 16, 'sub_arrival_location_standardized', self.bold)
        self.worksheet.write(self.row, sv_col + 17, 'sub_arrival_location_status', self.bold)
        self.worksheet.write(self.row, sv_col + 18, 'sub_arrival_date_as_source', self.bold)
        self.worksheet.write(self.row, sv_col + 19, 'sub_arrival_date_year', self.bold)
        self.worksheet.write(self.row, sv_col + 20, 'sub_arrival_date_month', self.bold)
        self.worksheet.write(self.row, sv_col + 21, 'sub_arrival_date_day', self.bold)
        self.worksheet.write(self.row, sv_col + 22, 'sub_arrival_date_year_to', self.bold)
        self.worksheet.write(self.row, sv_col + 23, 'sub_arrival_date_month_to', self.bold)
        self.worksheet.write(self.row, sv_col + 24, 'sub_arrival_date_day_to', self.bold)
        self.worksheet.write(self.row, sv_col + 25, 'sub_arrival_date_status', self.bold)
        self.worksheet.write(self.row, sv_col + 26, 'sub_arrival_date_relative', self.bold)
        self.worksheet.write(self.row, sv_col + 27, 'sub_vessel', self.bold)
        self.worksheet.write(self.row, sv_col + 28, 'sub_slaves', self.bold)
        self.worksheet.write(self.row, sv_col + 29, 'voyage_id', self.bold)
        self.worksheet.write(self.row, sv_col + 30, 'slaving_voyage_status', self.bold)
        self.worksheet.write(self.row, sv_col + 31, 'subvoyage_notes', self.bold)
        self.worksheet.write(self.row, sv_col + 32, 'sub_source', self.bold)

        sv_cursor = self.cnx.cursor(buffered=True)
        query = ("SELECT subvoyage_id, subvoyage_type, sub_dept_location, sub_dept_location_standardized, sub_dept_location_status, sub_dept_date_as_source, sub_dept_date_year, sub_dept_date_month, sub_dept_date_day, sub_dept_date_year_to, sub_dept_date_month_to, sub_dept_date_day_to, sub_dept_date_status, sub_dept_date_relative, sub_arrival_location, sub_arrival_location_standardized, sub_arrival_location_status, sub_arrival_date_as_source, sub_arrival_date_year, sub_arrival_date_month, sub_arrival_date_day, sub_arrival_date_year_to, sub_arrival_date_month_to, sub_arrival_date_day_to, sub_arrival_date_status, sub_arrival_date_relative, sub_vessel, sub_slaves,  voyage_id, slaving_voyage_status, subvoyage_notes, sub_source FROM subvoyage WHERE voyage_id = " + str(id))
        sv_cursor.execute(query)



        for (subvoyage_id, subvoyage_type, sub_dept_location, sub_dept_location_standardized, sub_dept_location_status, sub_dept_date_as_source, sub_dept_date_year, sub_dept_date_month, sub_dept_date_day, sub_dept_date_year_to, sub_dept_date_month_to, sub_dept_date_day_to, sub_dept_date_status, sub_dept_date_relative, sub_arrival_location, sub_arrival_location_standardized, sub_arrival_location_status, sub_arrival_date_as_source, sub_arrival_date_year, sub_arrival_date_month, sub_arrival_date_day, sub_arrival_date_year_to, sub_arrival_date_month_to, sub_arrival_date_day_to, sub_arrival_date_status, sub_arrival_date_relative, sub_vessel, sub_slaves,  voyage_id, slaving_voyage_status, subvoyage_notes, sub_source) in sv_cursor:
            self.row += 1
            self.worksheet.write(self.row, sv_col, subvoyage_id)
            self.worksheet.write(self.row, sv_col + 1, subvoyage_type)
            self.worksheet.write(self.row, sv_col + 2, sub_dept_location)
            self.worksheet.write(self.row, sv_col + 3, sub_dept_location_standardized)
            self.worksheet.write(self.row, sv_col + 4, sub_dept_location_status)
            self.worksheet.write(self.row, sv_col + 5, sub_dept_date_as_source)
            self.worksheet.write(self.row, sv_col + 6, sub_dept_date_as_source)
            self.worksheet.write(self.row, sv_col + 7, sub_dept_date_year)
            self.worksheet.write(self.row, sv_col + 8, sub_dept_date_month)
            self.worksheet.write(self.row, sv_col + 9, sub_dept_date_day)
            self.worksheet.write(self.row, sv_col + 10, sub_dept_date_year_to)
            self.worksheet.write(self.row, sv_col + 11, sub_dept_date_month_to)
            self.worksheet.write(self.row, sv_col + 12, sub_dept_date_day_to)
            self.worksheet.write(self.row, sv_col + 13, sub_dept_date_status)
            self.worksheet.write(self.row, sv_col + 14, sub_dept_date_relative)
            self.worksheet.write(self.row, sv_col + 15, sub_arrival_location)
            self.worksheet.write(self.row, sv_col + 16, sub_arrival_location_standardized)
            self.worksheet.write(self.row, sv_col + 17, sub_arrival_location_status)
            self.worksheet.write(self.row, sv_col + 18, sub_arrival_date_as_source)
            self.worksheet.write(self.row, sv_col + 19, sub_arrival_date_year)
            self.worksheet.write(self.row, sv_col + 20, sub_arrival_date_month)
            self.worksheet.write(self.row, sv_col + 21, sub_arrival_date_day)
            self.worksheet.write(self.row, sv_col + 22, sub_arrival_date_year_to)
            self.worksheet.write(self.row, sv_col + 23, sub_arrival_date_month_to)
            self.worksheet.write(self.row, sv_col + 24, sub_arrival_date_day_to)
            self.worksheet.write(self.row, sv_col + 25, sub_arrival_date_status)
            self.worksheet.write(self.row, sv_col + 26, sub_arrival_date_relative)
            self.worksheet.write(self.row, sv_col + 27, sub_vessel)
            self.worksheet.write(self.row, sv_col + 28, sub_slaves)
            self.worksheet.write(self.row, sv_col + 29, voyage_id)
            self.worksheet.write(self.row, sv_col + 30, slaving_voyage_status)
            self.worksheet.write(self.row, sv_col + 31, subvoyage_notes)
            self.worksheet.write(self.row, sv_col + 32, sub_source)

    def make_dump(self):
        self.workbook = xlsxwriter.Workbook('static/esta.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.col = 0
        self.row = 0

        self.cnx = mysql.connector.connect(user='root', password='bonzo', host='127.0.0.1', database='esta_live')
        self.cursor = self.cnx.cursor(buffered=True)

        query = ("SELECT voyage_id, summary, year, DATE_FORMAT(last_mutation, \"%d-%m-%Y\") AS last_mutation FROM voyage")
        self.cursor.execute(query)

        self.bold = self.workbook.add_format({'bold': True})
        self.worksheet.write(self.row, self.col, 'voyage_id', self.bold)
        self.worksheet.write(self.row, self.col + 1, 'summary', self.bold)
        self.worksheet.write(self.row, self.col + 2, 'year', self.bold)
        self.worksheet.write(self.row, self.col + 3, 'last_mutation', self.bold)

        for (voyage_id, summary, year, last_mutation) in self.cursor:
            self.row += 1
            self.worksheet.write(self.row, self.col, voyage_id)
            self.worksheet.write(self.row, self.col + 1, summary)
            self.worksheet.write(self.row, self.col + 2, year)
            self.worksheet.write(self.row, self.col + 3, last_mutation)
            self.write_subvoyage(voyage_id)
            self.row += 1

        self.cnx.close()
        self.workbook.close()
