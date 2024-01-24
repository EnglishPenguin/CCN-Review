from classes import *
from tkinter import simpledialog as sd
from tkinter import messagebox as mb
from logger_setup import logger

if __name__ == '__main__':
    while True:
        output_process = sd.askinteger("Output Review Process", "Specify File Review or Email Prep by picking a number:\n\n1 - File Review \n2 - Email Prep \n\n Press 'Cancel' to Exit", minvalue=1, maxvalue=2)
        if output_process == 1:
            ccn_review = Output_Review()
            ccn_review.get_file_date()
            ccn_review.convert_dates_to_strings()
            ccn_review.ask_if_correct_date()
            ccn_review.move_templates_to_detination()
            ccn_review.create_data_frames()
            ccn_review.populate_stat_comm_cat_act()
            ccn_review.get_rep_submissions()
            ccn_review.prep_and_export_file()
        elif output_process == 2:
            email_prep = Email_Prep()
            email_prep.get_file_date()
            email_prep.convert_dates_to_strings()
            email_prep.ask_if_correct_date()
            email_prep.create_email_dataframes()
            email_prep.get_users_to_email()
            email_prep.drop_columns_reorder_write_to_file()
            email_prep.prep_email_list()
            email_prep.create_email()
        elif output_process == None:
            mb.showinfo("Closing", "Thank you, exiting program")
            logger.info("user ending program")
            exit()
        else:
            mb.showerror("Invalid Input", "Please make a valid selection")
