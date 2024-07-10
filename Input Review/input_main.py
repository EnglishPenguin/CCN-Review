from input_classes import *
from tkinter import simpledialog as sd
from tkinter import messagebox as mb
from input_logger_setup import logger

def main():
    dt_func = Date_Functions()
    dt_func.run()
    while True:
        input_process = sd.askinteger(
            "Input Review Process", 
            "Specify Input Combine, Input Review, or File to CSV by picking a number:\n\n1 - Input Combine \n2 - Input Review \n3 - File to CSV \n\n Press 'Cancel' to Exit", 
            minvalue=1, 
            maxvalue=3
        )
        if input_process == 1:
            logger.info('starting input combine')
            fc = File_Combine(
                dt_func.fd_mmddyyy, 
                dt_func.fd_mm, 
                dt_func.fd_yyyy, 
                dt_func.fd_spaces, 
                dt_func.fd_slashes
                )
            fc.run()
        elif input_process == 2:
            logger.info('starting input review')
            ir = Input_Review(
                dt_func.fd_mmddyyy, 
                dt_func.fd_mm, 
                dt_func.fd_yyyy, 
                dt_func.fd_spaces, 
                dt_func.fd_slashes
                )
            ir.run()
        elif input_process == 3:
            logger.info('starting file to csv')
            f2c = File_To_CSV(
                dt_func.fd_mmddyyy, 
                dt_func.fd_mm, 
                dt_func.fd_yyyy, 
                dt_func.fd_spaces, 
                dt_func.fd_slashes
            )
            f2c.run()
        elif input_process == None:
            logger.info('user exiting program')
            exit()

main()