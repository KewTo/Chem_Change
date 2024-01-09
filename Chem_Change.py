# import pygetwindow as gw

WIINGS = input("Copy and paste the WIINGS order #: ")
num_of_lines = int(input("How many lines are you changing?: "))

while True:
    try:
        for _ in range(num_of_lines):
            while True:
                COAT_TOOL = input("Which tool are you changing: ")  # ACT01/CCT01/RCT03/RCT05/GCT01
                COAT_TOOL = COAT_TOOL.upper()

                if COAT_TOOL == 'ACT01' or COAT_TOOL == 'CCT01':
                    CHEM_LINE = input("Which line are you changing: ")
                    CHEM_LINE = CHEM_LINE.upper()
                    if CHEM_LINE == 'ACT12' or CHEM_LINE == 'ACT13' or CHEM_LINE == 'C21' or CHEM_LINE == 'C21B':
                        BOTTLE = 'EMX-7306'
                    elif CHEM_LINE == 'ACT11' or CHEM_LINE == 'ACT14' or CHEM_LINE == 'C14':
                        BOTTLE = 'EMX-7310'
                    EXP_DATE = input("Input the expiration date: ")
                    BOTTLE_TIME = input("Input time bottle was put on warming shelf: ")
                    BATCH = input("Input the batch number: ")
                    BATCH = BATCH.upper()
                    BUDDY = input("Input your buddy name: ")
                    BUDDY = BUDDY.upper()

                    information = {
                        'TOOL': COAT_TOOL,
                        'CHEMICAL_LINE': CHEM_LINE,
                        'BATCH_NUMBER': BATCH,
                        'EXPIRATION_DATE': EXP_DATE,
                        'WIINGS_NUMBER': WIINGS,
                        'BUDDY_NAME': BUDDY,
                        'BOTTLE_FROM_WARMING_SHELF': BOTTLE_TIME,
                        'BOTTLE_NAME': BOTTLE
                    }

                    print(information)

                elif COAT_TOOL == 'RCT03' or COAT_TOOL == 'RCT05' or COAT_TOOL == 'GCT01':
                    CHEM_LINE = input("Which line are you changing: ")
                    CHEM_LINE = CHEM_LINE.upper()
                    if CHEM_LINE == 'PCAR' or CHEM_LINE == 'PCARB':
                        BOTTLE = 'FEP-171'
                        viscosity = '3mPa*s'
                    elif CHEM_LINE == '8K':
                        BOTTLE = 'FEP-171'
                        viscosity = '6mPa*s'
                    elif CHEM_LINE == 'P53' or CHEM_LINE == 'P53B' or CHEM_LINE == 'P53C':
                        BOTTLE = 'MES EP500JE 1.9X'
                    elif CHEM_LINE == 'P37' or CHEM_LINE == 'P37B':
                        BOTTLE = 'SEBP-9012'
                    elif CHEM_LINE == 'P23':
                        BOTTLE = 'OEBR-CAP130T4'
                    elif CHEM_LINE == 'P81':
                        BOTTLE = 'MES EP561JE-1.9cP'
                    elif CHEM_LINE == 'P94' or CHEM_LINE == 'P94B':
                        BOTTLE = 'SEBP-602F'
                    elif CHEM_LINE == 'BARC':
                        BOTTLE = 'E2 STACK AL9412-302'
                    elif CHEM_LINE == 'N19' or CHEM_LINE == 'N19B':
                        BOTTLE = 'SEBN-306F'
                    else:
                        BOTTLE = ''
                    BATCH = input("Input the batch number: ")
                    BATCH = BATCH.upper()
                    EXP_DATE = input("Input the expiration date: ")
                    BUDDY = input("Input your buddy name: ")
                    BUDDY = BUDDY.upper()

                    information = {
                        'TOOL': COAT_TOOL,
                        'CHEMICAL_LINE': CHEM_LINE,
                        'BATCH_NUMBER': BATCH,
                        'EXPIRATION_DATE': EXP_DATE,
                        'WIINGS_NUMBER': WIINGS,
                        'BUDDY_NAME': BUDDY,
                        'BOTTLE_NAME': BOTTLE
                    }

                else:
                    print('Invalid Tool; please enter a valid tool')

    except ValueError:
        print('Invalid Tool; please enter a valid tool')
        continue

    else:
        break


def prompt_script():
    TOOL, CHEMICAL_LINE, BATCH_NUMBER, EXPIRATION_DATE, WIINGS_NUMBER, BUDDY_NAME, BOTTLE_FROM_WARMING_SHELF, BOTTLE_NAME = tool_prompt().values()

    if TOOL == 'ACT01' or TOOL == 'CCT01':
        print(
            f" {CHEMICAL_LINE} REPLACED EMPTY BOTTLE WITH BOTTLE FROM WARMING SHELF \n  "
            f"\n "
            f"BOTTLE FROM WARMING SHELF: {BOTTLE_FROM_WARMING_SHELF} \n "
            f"\n "
            f"{BOTTLE_NAME} \n "
            f"BATCH: {BATCH_NUMBER} \n "
            f"EXP DATE: {EXPIRATION_DATE} \n "
            f"WIINGS: {WIINGS_NUMBER} \n "
            f"BUDDY: {BUDDY_NAME} \n "
            f"\n "
            f"KT "
        )

    elif TOOL == 'RCT03' or TOOL == 'RCT05' or TOOL == 'GCT01':
        print(
            f" {TOOL}-{CHEMICAL_LINE} REPLACED EMPTY BOTTLE WITH BOTTLE FROM FRIDGE \n  "
            f"\n "
            f"{BOTTLE_NAME} \n "
            f"BATCH: {BATCH_NUMBER} \n "
            f"EXP DATE: {EXPIRATION_DATE} \n "
            f"WIINGS: {WIINGS_NUMBER} \n "
            f"BUDDY: {BUDDY_NAME} \n "
            f"\n "
            f"KT "
        )


def get_Workstream():
    win = gw.getWindowsWithTitle('EXTRA! X-treme')[0]
    win.activate()


#tool_prompt()

if __name__ == '__main__':
    pass
