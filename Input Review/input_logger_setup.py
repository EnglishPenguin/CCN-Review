from loguru import logger

logger.add("M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/References/Log Files/Input Review/{time:YYYY-MM-DD}.log",
           format="{time:YYYY-MM-DD at HH:mm:ss} | {level} - {message}",
           colorize=True, backtrace=True, diagnose=True, level='DEBUG', retention='90 days')
