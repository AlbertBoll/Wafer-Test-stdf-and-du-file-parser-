import copy
import csv
import os.path
import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor
from fpdf.enums import XPos, YPos
import numpy as np
from stdf import *
import itertools
import math
import warnings

# warnings.filterwarnings('ignore', category=RuntimeWarning)
# warnings.filterwarnings('ignore', category=DeprecationWarning)

plt.switch_backend('agg')


def split_dict(input_dict, num_parts=3):
    dict_len = len(input_dict)
    chunk_size = math.ceil(dict_len / num_parts)

    items = list(input_dict.items())

    return [dict(items[i:i + chunk_size]) for i in range(0, dict_len, chunk_size)]


def wafer_map_task(tests_dict, x, y, y_shift, wafer_dir=""):
    for test_id, test_info in tests_dict.items():
        title_ = f'Test ID: {test_id} Wafer Map'
        generate_wafer_map(x, y, y_shift, test_info.b_passed, title_,
                           str(test_id) + "_before_wafermap.png",
                           wafer_dir)


def wafer_plot_task(filter_type, tests_dict, fig, ax1, ax2, directory):
    trend_title = "Trend Plot"
    histogram = "Histogram"
    plt.style.use('ggplot')

    for test_id, test_info in tests_dict.items():
        if filter_type == FilterType.ALL_RUN:
            test_info.dropped = 0
        mean = round(test_info.stats.mean, 3)
        median = round(test_info.stats.median, 3)
        std = round(test_info.stats.std, 3)
        cpk = round(test_info.stats.cpk, 3)
        unit = test_info.unit
        l_lim = test_info.lower_limit
        h_lim = test_info.upper_limit
        number_points = len(test_info.values)
        number_array = np.arange(1, number_points + 1)
        value_array = test_info.values
        ax1.scatter(number_array, value_array, c="blue", s=2, marker="o")
        ax1.plot(number_array, value_array, "r-", linewidth=1)
        ax1.plot(number_array, np.repeat(h_lim, number_points), linewidth=2, linestyle="--", label=f"HLIM: {h_lim}")
        ax1.plot(number_array, np.repeat(l_lim, number_points), linewidth=2, linestyle="--", label=f"LLIM: {l_lim}")
        ax1.legend(loc="upper right")
        ax1.set_xlabel("runs")
        limit_gap = ((h_lim - l_lim) / 5.) * 0.2
        ax1.set_ylabel(unit)
        ax1.set_ylim(l_lim - limit_gap, h_lim + limit_gap)
        ax1.set_title(trend_title)

        ax2.axvline(x=l_lim, color="blue", linestyle="-", label=f"ll={l_lim}")
        ax2.axvline(x=h_lim, color="blue", linestyle="-", label=f"ul={h_lim}")
        ax2.axvline(x=median, color="grey", linestyle="--", label=f"median={median}")

        ax2.axvline(x=mean - std, color="green", linestyle="-",
                    label=f"{from_enum_to_string(FilterType.ONE_SIGMA)} ll={round(mean - std, 3)}")
        ax2.axvline(x=mean + std, color="green", linestyle="-",
                    label=f"{from_enum_to_string(FilterType.ONE_SIGMA)} ul={round(mean + std, 3)}")

        ax2.hist(value_array, bins=10, color="skyblue", edgecolor="black", align='mid')

        ax2.set_xlabel(unit)
        ax2.tick_params(axis='x', labelrotation=45)
        ax2.xaxis.set_major_formatter(FormatStrFormatter('%.2f'))
        ax2.set_title(histogram)
        ax2.legend(loc="upper right")
        plt.tight_layout()
        fig.suptitle(f"Test {test_id} - {test_info.test_description}\n"
                     f"fallout = {test_info.dropped} out of {test_info.total_run} "
                     f"({round((test_info.dropped / test_info.total_run) * 100, 2)}%)\n\u03BC = {mean} \u03C3 = {std} cpk = {round(cpk, 1)}\n"
                     f"dropped = {test_info.dropped}",
                     fontsize=14)

        plot_file_name = directory + str(test_id) + ".png"
        fig.savefig(plot_file_name)
        ax1.cla()
        ax2.cla()


def task_1(filter_type, tests_dict, x, y, y_shift, fig, ax1, ax2, directory):
    trend_title = "Trend Plot"
    histogram = "Histogram"
    title = from_enum_to_string(filter_type)
    plt.style.use('ggplot')
    for test_id, test_info in tests_dict.items():
        title_ = f'Test ID: {test_id} Wafer Map'
        generate_wafer_map(x, y, y_shift, test_info.b_passed, title_,
                           str(test_id) + "_before_wafermap.png",
                           title + " Wafer Map/")

        if filter_type == FilterType.ALL_RUN:
            test_info.dropped = 0
        mean = round(test_info.stats.mean, 3)
        median = round(test_info.stats.median, 3)
        std = round(test_info.stats.std, 3)
        cpk = round(test_info.stats.cpk, 3)
        unit = test_info.unit
        l_lim = test_info.lower_limit
        h_lim = test_info.upper_limit
        number_points = len(test_info.values)
        number_array = np.arange(1, number_points + 1)
        value_array = test_info.values
        ax1.scatter(number_array, value_array, c="blue", s=2, marker="o")
        ax1.plot(number_array, value_array, "r-", linewidth=1)
        ax1.plot(number_array, np.repeat(h_lim, number_points), linewidth=2, linestyle="--", label=f"HLIM: {h_lim}")
        ax1.plot(number_array, np.repeat(l_lim, number_points), linewidth=2, linestyle="--", label=f"LLIM: {l_lim}")
        ax1.legend(loc="upper right")
        ax1.set_xlabel("runs")
        limit_gap = ((h_lim - l_lim) / 5.) * 0.2
        ax1.set_ylabel(unit)
        ax1.set_ylim(l_lim - limit_gap, h_lim + limit_gap)
        ax1.set_title(trend_title)

        ax2.axvline(x=l_lim, color="blue", linestyle="-", label=f"ll={l_lim}")
        ax2.axvline(x=h_lim, color="blue", linestyle="-", label=f"ul={h_lim}")
        ax2.axvline(x=median, color="grey", linestyle="--", label=f"median={median}")

        ax2.axvline(x=mean - std, color="green", linestyle="-",
                    label=f"{from_enum_to_string(FilterType.ONE_SIGMA)} ll={round(mean - std, 3)}")
        ax2.axvline(x=mean + std, color="green", linestyle="-",
                    label=f"{from_enum_to_string(FilterType.ONE_SIGMA)} ul={round(mean + std, 3)}")

        ax2.hist(value_array, bins=10, color="skyblue", edgecolor="black", align='mid')

        ax2.set_xlabel(unit)
        ax2.tick_params(axis='x', labelrotation=45)
        ax2.xaxis.set_major_formatter(FormatStrFormatter('%.2f'))
        ax2.set_title(histogram)
        ax2.legend(loc="upper right")
        plt.tight_layout()
        # fig.suptitle(str(test_id) + ": " + test_info.test_description, fontsize="x-large")
        fig.suptitle(f"Test {test_id} - {test_info.test_description}\n"
                     f"fallout = {test_info.dropped} out of {test_info.total_run} "
                     f"({round((test_info.dropped / test_info.total_run) * 100, 2)}%)\n\u03BC = {mean} \u03C3 = {std} cpk = {round(cpk, 1)}\n"
                     f"dropped = {test_info.dropped}",
                     fontsize=14)

        # temp = ""
        # if filter_type == FilterType.PASSED_RUN:
        #     temp = "Passed Runs"
        # elif filter_type == FilterType.ALL_RUN:
        #     temp = "All Runs"
        #
        # directory = temp + " Plots/"
        # if not os.path.exists(directory):
        #     os.makedirs(directory, exist_ok=True)

        # wafer_map_file_name = f"{title}_wafer_map/{test_id}_before_after_wafermap.png"
        plot_file_name = directory + str(test_id) + ".png"
        fig.savefig(plot_file_name)
        ax1.cla()
        ax2.cla()


class WaferRunInfoDU:
    def __init__(self):
        self.family_id = ""
        self.product_id = ""
        self.start_time = ""
        self.lot = ""

    def __str__(self):
        return f"family id: {self.family_id}\nproduct id: {self.product_id}\nstart time: {self.start_time}\nlot: {self.lot}"

    def __repr__(self):
        return f"family id: {self.family_id}\nproduct id: {self.product_id}\nstart time: {self.start_time}\nlot: {self.lot}"

    def AllDataFilledIn(self):
        return len(self.family_id) > 0 and len(self.product_id) > 0 and len(self.start_time) > 0 and len(self.lot) > 0


class TestFieldAndResult:
    def __init__(self):
        self.test_description = ""
        self.lower_limit = 0
        self.upper_limit = 0
        self.unit = ""
        self.value = 0
        self.b_passed = None


class DieRunResult:
    def __init__(self):
        self.each_test_dict = {}
        self.x = 0
        self.y = 0
        self.b_passed = None


class TestResult:
    def __init__(self):
        self.test_description = ""
        self.lower_limit = 0.0
        self.upper_limit = 0.0
        self.unit = ""
        self.b_passed = np.array([], dtype=object)
        self.values = np.array([], dtype=np.float64)
        self.stats = StatisticalResults()
        self.total_run = 0
        self.dropped = 0

    def generate_statistics(self):

        if len(self.values) > 1:
            self.stats.min = np.min(self.values)
            self.stats.max = np.max(self.values)
            self.stats.mean = np.mean(self.values)
            self.stats.std = 0.0 if np.std(self.values) <= 1.0e-6 else np.std(self.values)
            self.stats.range = self.stats.max - self.stats.min
            self.stats.median = np.median(self.values)

            try:
                self.stats.cp = (self.upper_limit - self.lower_limit) / (6. * self.stats.std)
                self.stats.cpk = min((self.upper_limit - self.stats.mean) / (3 * self.stats.std),
                                     (self.stats.mean - self.lower_limit) / (3 * self.stats.std))
            except:
                self.stats.cp = 9999.999
                self.stats.cpk = 9999.999
        elif len(self.values) == 1:
            self.stats.min = self.values[0]
            self.stats.max = self.values[0]
            self.stats.mean = self.values[0]
            self.stats.std = 0.0
            self.stats.range = self.stats.max - self.stats.min
            self.stats.median = self.values[0]
            self.stats.cp = 9999.999
            self.stats.cpk = 9999.999


class WaferRunCSVParser(object):
    def __init__(self):
        # self.file_name = file_name
        self.b_loaded = False
        self.dies_run_result = []
        self.sorted_cpk_list = []
        self.wafer_run_info = WaferRunInfoDU()
        self.test_id_list = []
        self.test_description_list = []
        self.tests_dict = {}
        self.outlier_excluded_tests_dict = {}

    def transform_data(self):
        x_list = [element.x for element in self.dies_run_result]
        y_list = [element.y for element in self.dies_run_result]
        x_min = min(x_list)
        self.x = np.array(x_list, dtype=np.int16) + abs(x_min)
        self.y = np.array(y_list, dtype=np.int16)
        self.b_wafers_passed = np.array([element.b_passed for element in self.dies_run_result], dtype=bool)

        for test_id in self.test_id_list:
            for die in self.dies_run_result:
                if test_id not in die.each_test_dict:
                    die.each_test_dict[test_id] = None

        # print(self.test_id_list)
        for die in self.dies_run_result:
            for _id, run in die.each_test_dict.items():
                if _id not in self.tests_dict:
                    self.tests_dict[_id] = TestResult()
                if run is not None:
                    self.tests_dict[_id].test_description = run.test_description
                    self.tests_dict[_id].lower_limit = run.lower_limit
                    self.tests_dict[_id].upper_limit = run.upper_limit
                    self.tests_dict[_id].unit = run.unit
                    self.tests_dict[_id].b_passed = np.append(self.tests_dict[_id].b_passed, run.b_passed)
                    self.tests_dict[_id].values = np.append(self.tests_dict[_id].values, run.value)
                else:
                    self.tests_dict[_id].b_passed = np.append(self.tests_dict[_id].b_passed, None)

        for key, each_test_for_all_runs in self.tests_dict.items():
            self.test_description_list.append(each_test_for_all_runs.test_description)
            each_test_for_all_runs.total_run = len(each_test_for_all_runs.values)
            arr_cleaned = np.where(each_test_for_all_runs.b_passed == None, 0, each_test_for_all_runs.b_passed)
            filter_none_array = each_test_for_all_runs.b_passed[
                each_test_for_all_runs.b_passed != np.array(None)].astype(bool)
            test_result = copy.deepcopy(each_test_for_all_runs)
            test_result.values = each_test_for_all_runs.values[filter_none_array]
            test_result.generate_statistics()
            each_test_for_all_runs.dropped = each_test_for_all_runs.total_run - np.sum(arr_cleaned)
            test_result.dropped = each_test_for_all_runs.dropped
            self.outlier_excluded_tests_dict[key] = test_result
            each_test_for_all_runs.generate_statistics()
        return self

    def parse(self, file_name):
        need_collect_test_id = True
        with open(file_name, 'r') as file:
            csv_reader = csv.reader(file)
            for row in csv_reader:
                # print(row)
                if not self.wafer_run_info.AllDataFilledIn():
                    if len(row[0]) >= 5:
                        # fetch the family id
                        if row[0][0:3] == "FAM":
                            self.wafer_run_info.family_id = row[0][4:]
                        # fetch the product id
                        elif row[0][0:2] == "PN":
                            self.wafer_run_info.product_id = row[0][3:]
                        # fetch the start data
                        elif row[0][0:2] == "SD":
                            self.wafer_run_info.start_time = row[0][3:]
                        # fetch the start time
                        elif row[0][0:2] == "ST":
                            self.wafer_run_info.start_time += (" " + row[0][3:])
                        # fetch the lot
                        elif row[0][0:2] == "SN":
                            self.wafer_run_info.lot = row[0][3:9]

                if len(row[0]) > 4 and row[0][0:4] == "WAFX":
                    die_result = DieRunResult()
                    die_result.x = int(row[0][5:])
                    wafer_y_row = next(csv.reader(itertools.islice(file, 0, 1)))
                    die_result.y = int(wafer_y_row[0][5:])
                    begin_result_row = next(csv.reader(itertools.islice(file, 1, 2)))

                    _id = 0
                    while begin_result_row[0] != "END TRAILER":
                        begin_result_row = next(csv.reader(itertools.islice(file, 0, 1)))
                        test_field_result = TestFieldAndResult()
                        if len(begin_result_row) > 1:
                            if begin_result_row[-4] != '':
                                _id = int(begin_result_row[0])
                                if _id not in self.test_id_list:
                                    self.test_id_list.append(_id)
                                description = begin_result_row[1]
                                unit = begin_result_row[5]
                                if unit.lower().strip() == "hex":
                                    lower_limit = float(int(begin_result_row[2], 16))
                                    upper_limit = float(int(begin_result_row[3], 16))
                                    value = float(int(begin_result_row[4], 16))
                                else:
                                    lower_limit = float(begin_result_row[2])
                                    upper_limit = float(begin_result_row[3])
                                    value = float(begin_result_row[4])
                                b_passed = True if begin_result_row[-1] == 'P' else False

                                test_field_result.test_description = description
                                test_field_result.b_passed = b_passed
                                test_field_result.value = value
                                test_field_result.lower_limit = lower_limit
                                test_field_result.upper_limit = upper_limit
                                test_field_result.unit = unit
                        else:
                            if len(begin_result_row[0]) >= 4:
                                if begin_result_row[0][0:3] == "RST":
                                    if begin_result_row[0][4:] == "FAIL":
                                        die_result.b_passed = False
                                    else:
                                        die_result.b_passed = True
                        if _id not in die_result.each_test_dict:
                            die_result.each_test_dict[_id] = test_field_result
                    self.dies_run_result.append(die_result)

        return self

    def convert_to_stats_report_pdfs(self, generation_list, filter_type, progress_bar, generate_button,
                                     progress_label=None, src_dict={}):
        for filename in generation_list:
            print(filename)
            self.convert_to_stats_report_pdf(filename, filter_type, progress_bar, progress_label, src_dict)
            self.b_loaded = False
            self.wafer_run_info = WaferRunInfoDU()
            self.test_id_list.clear()
            self.test_description_list.clear()
            self.tests_dict.clear()
            self.outlier_excluded_tests_dict.clear()
            self.dies_run_result.clear()
            self.sorted_cpk_list.clear()

        generate_button.state(["!disabled"])  # active the button again

    def convert_to_stats_report_pdf(self, file_name, filter_type, progress_bar, progress_label=None, src_dict={}):

        with ProcessPoolExecutor() as executor:
            if not self.b_loaded:
                self.parse(file_name).transform_data()
                self.b_loaded = True
            plt.style.use('ggplot')
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 8), constrained_layout=True)
            tests_dict = None
            file_name = os.path.splitext(os.path.basename(file_name))[0]
            classification = ""
            if filter_type == FilterType.PASSED_RUN:
                classification = "Filtered"
            elif filter_type == FilterType.ALL_RUN:
                classification = ""

            plot_directory = file_name + "_" + self.wafer_run_info.family_id + "_" + classification + " Plots/"
            test_wafer_directory = file_name + "_" + classification + " Wafer Map/"
            whole_wafer_directory = file_name + "_" + self.wafer_run_info.family_id + " Wafer Map/"

            if filter_type == FilterType.ALL_RUN:
                tests_dict = copy.deepcopy(self.tests_dict)
            elif filter_type == FilterType.PASSED_RUN:
                tests_dict = copy.deepcopy(self.outlier_excluded_tests_dict)

            y_shift = max(self.y)
            x_shift = max(self.x) - min(self.x)
            shift = max(y_shift, x_shift)
            title = from_enum_to_string(filter_type)

            # os.makedirs(output_file_dir, exist_ok=True)
            os.makedirs(plot_directory, exist_ok=True)
            os.makedirs(test_wafer_directory, exist_ok=True)
            os.makedirs(whole_wafer_directory, exist_ok=True)

            # generate whole wafer map first
            generate_wafer_map(self.x, self.y, shift, self.b_wafers_passed,
                               f'{self.wafer_run_info.family_id} Wafer Run',
                               f'{file_name}_{self.wafer_run_info.family_id}_Wafer_Run.png',
                               whole_wafer_directory)

            # start = time.time()
            # generate
            split_dict_list = split_dict(tests_dict, 6)

            # with ProcessPoolExecutor() as executor:
            wafer_map_futures = [executor.submit(wafer_map_task, _dict, self.x, self.y, shift,
                                                 test_wafer_directory) for _dict in split_dict_list]
            plot_futures = [executor.submit(wafer_plot_task, filter_type, dict_, fig, ax1, ax2, plot_directory)
                            for dict_ in split_dict_list]

            for map_future, plot_future in zip(wafer_map_futures, plot_futures):
                map_future.result()
                plot_future.result()

            cpk_first_n = None
            # pick first 40 smallest cpk test
            if filter_type == FilterType.ALL_RUN:
                sorted_ = {k: v for k, v in sorted(self.tests_dict.items(), key=lambda item: item[1].stats.cpk)}
                cpk_first_n = list(sorted_.items())[:40]
            elif filter_type == FilterType.PASSED_RUN:
                sorted_ = {k: v for k, v in sorted(self.outlier_excluded_tests_dict.items(), key=lambda item: item[1].stats.cpk)}
                cpk_first_n = list(sorted_.items())[:40]

            # generate pdf stats report
            statistics_report = "Statistics Report"
            pdf = FPDF()

            # add title
            pdf.add_page()
            pdf.set_font("helvetica", size=40, style="IB")
            pdf.set_text_color(255, 0, 0)
            pdf.cell(195, 70, txt=statistics_report, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

            pdf.set_font("helvetica", size=20, style="IB")
            pdf.set_text_color(0, 0, 250)
            pdf.cell(195, 10, txt=self.wafer_run_info.family_id, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font("helvetica", size=16, style="B")
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 10, txt=str(" " * 40 + str(self.wafer_run_info.product_id)),
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align="L")

            pdf.set_font("helvetica", size=16, style="B")
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 10, txt=" " * 40 + self.wafer_run_info.start_time,
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align="L")

            pdf.set_font("helvetica", size=16, style="B")
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 10, txt=" " * 96 + "Lot: " + str(self.wafer_run_info.lot),
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align="C")

            total_runs = len(self.dies_run_result)
            passed_run = np.sum(self.b_wafers_passed)
            yield_rate = passed_run / total_runs * 100

            pdf.set_font("helvetica", size=16, style="B")
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 10, txt=" " * 101 + "All Runs: "+str(total_runs),
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align="C")

            pdf.set_font("helvetica", size=16, style="B")
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 10, txt=" " * 107 + "Passed Runs: " + str(passed_run),
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align="C")

            pdf.set_font("helvetica", size=16, style="B")
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 10, txt=" " * 97 + "Yield: " + str(round(yield_rate, 1))+"%",
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, align="C")

            pdf.image(f'{file_name}_{self.wafer_run_info.family_id} Wafer Map/{file_name}_{self.wafer_run_info.family_id}_Wafer_Run.png',
                      x=45, y=165, w=120, h=110)

            table_data = [["TNUM", "SITE", "RANGE", "STD", "CP", "CPK", "UNITS", "TNAME"]]
            link_dic = {}
            for k, v in cpk_first_n:
                temp = []
                temp.append(str(k))
                temp.append(1)
                temp.append(round(v.stats.range, 3))
                temp.append(round(v.stats.std, 3))
                temp.append(round(v.stats.cp, 3))
                temp.append(round(v.stats.cpk, 3))
                temp.append(v.unit)
                temp.append(v.test_description)
                link_dic[v.test_description] = pdf.add_link()
                table_data.append(temp)

            pdf.add_page()
            pdf.set_font("helvetica", size=20, style="B")
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 10, txt=" " * 85 + "Cpk Summary", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
            #
            pdf.ln(5)
            pdf.set_font('Helvetica', 'b', 8)
            pdf.set_text_color(0, 0, 0)

            for row_index, data_row in enumerate(table_data):
                for column_index, data in enumerate(data_row):
                    if column_index != 7:
                        if row_index == 0 or column_index != 0:
                            pdf.cell(w=16, h=6, txt=" "*5 + str(data), border=1, ln=0, align="R")
                        else:
                            pdf.cell(w=16, h=6, txt=" "*5 + str(data), border=1, ln=0, align="R", link=link_dic[data_row[7]])
                    else:
                        if row_index == 0:
                            pdf.cell(w=79, h=6, txt=" "*5 + str(data), border=1, ln=0, align="R")
                        else:
                            pdf.cell(w=79, h=6, txt=" "*5 + str(data), border=1, ln=0, align="R", link=link_dic[data])
                pdf.ln()

            # add trend plot and histogram
            index = 0
            load = len(self.tests_dict)
            for test_id, test_info in tests_dict.items():
                pdf.add_page()
                test_description = test_info.test_description
                if test_description in link_dic:
                    pdf.set_link(link_dic[test_description])
                min_ = round(test_info.stats.min, 3)
                max_ = round(test_info.stats.max, 3)
                mean = round(test_info.stats.mean, 3)
                # median = round(test_info.stats.median, 3)
                range_ = round(test_info.stats.range, 3)
                std = round(test_info.stats.std, 3)
                cp = round(test_info.stats.cp, 3)
                cpk = round(test_info.stats.cpk, 3)
                unit = test_info.unit
                # l_lim = test_info.lower_limit
                # h_lim = test_info.upper_limit
                pdf.add_font(fname='DejaVuSansCondensed.ttf')
                pdf.set_font("helvetica", style="B", size=15)
                pdf.set_text_color(0, 0, 0)
                pdf.cell(180, 10, txt=f"Test: {test_id}", ln=1, align="C")
                pdf.set_font("helvetica", style="B", size=20)
                pdf.set_text_color(0, 0, 0)
                pdf.cell(180, 10, txt=f"{test_description}", ln=1, align="C")
                pdf.ln(3)
                pdf.set_font("DejaVuSansCondensed", size=10)
                pdf.set_text_color(0, 0, 0)
                parameter_table = [["SITE", "FILTER", "MIN", "MEAN", "MAX", "RANGE", "STD", "CP", "CPK", "UNITS"],
                                   [1, title, min_, mean, max_, range_, std, cp, cpk, unit]]

                for data_row in parameter_table:
                    for row_index, data in enumerate(data_row):
                        pdf.cell(w=19, h=6, txt=" " * 4 + str(data), border=1, ln=0, align="R")
                    pdf.ln()
                plot_file_path = plot_directory + str(test_id) + ".png"
                # insert wafer map
                pdf.image(f'{test_wafer_directory}{test_id}_before_wafermap.png', x=45, y=55, w=120, h=110)


                # insert plot
                # print(plot_file_path)
                pdf.image(plot_file_path, x=0, y=170, w=200, h=100)
                index += 1
                progress_val = index / float(load) * 100
                progress_bar["value"] = progress_val
                if progress_label is not None:
                    progress_label.config(text=f"{progress_val:.0f}%")
                progress_bar.update_idletasks()

            file_name_output = os.path.join(os.getcwd(), file_name + ".pdf")
            pdf.output(file_name_output)

            # add outline of pdf
            pdf_ = Pdf.open(file_name_output)
            outline_list = []
            with pdf_.open_outline() as outline:
                oi = OutlineItem("Cover page", 0)
                outline_list.append(oi)
                oi = OutlineItem("Cpk summary", 1)
                outline_list.append(oi)
                page_index = 1
                for test_id, test_info in tests_dict.items():
                    page_index += 1
                    oi = OutlineItem(str(test_id) + ": " + test_info.test_description, page_index)
                    outline_list.append(oi)
                outline.root.extend(outline_list)

            # output pdf to current working directory/statistics report/
            if not os.path.exists("statistics report/"):
                os.makedirs("statistics report/", exist_ok=True)
            pdf_.save(os.getcwd() + "/statistics report/" + file_name + "_" +
                      self.wafer_run_info.family_id + "_" + classification + ".pdf")

            if "src_path" not in src_dict:
                src_dict["src_path"] = []
            src_dict["src_path"].append(file_name_output)
            threading.Thread(target=self.show_message_non_blocking,
                             args=("Completion!", f"{file_name}_{self.wafer_run_info.family_id} "
                                                  f""f"wafer run stats report successfully!")).start()
            progress_bar["value"] = 0
            if progress_label is not None:
                progress_label.config(text=f"0%")
            progress_bar.update_idletasks()

            self.remove_directory(plot_directory)
            self.remove_directory(test_wafer_directory)
            self.remove_directory(whole_wafer_directory)

            # generate_button.state(["!disabled"])  # active the button again

        return self

    def show_message_non_blocking(self, title, message):
        messagebox.showinfo(title, message)

    @staticmethod
    def remove_directory(directory_path):
        if os.path.exists(directory_path):
            try:
                shutil.rmtree(directory_path)
                print(f"Directory '{directory_path}' and its contents have been removed.")
            except OSError as e:
                print(f"Error: {e} - Could not remove directory '{directory_path}'.")
        else:
            print(f"Directory '{directory_path}' does not exist.")






