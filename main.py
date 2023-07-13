from scripts.email_handler import EmailHandler
from datetime import date
import time
import scripts.sap_bot as sap_robot
import scripts.transformadados as tf
import scripts.os_handler as osha
import os
import scripts.logger as logger


ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
user = os.environ.get('USERNAME')
exception = ['mmc',
            'python',
            'OUTLOOK',
            'Code',
            ]

def main():
    osha.close_running_apps(exception)
    sr = sap_robot.SapBot()
    sr.login()
    tf.routine(sr)
    sr.close_sap()
    tf.pandas_cleaning()
    try:
        with EmailHandler() as mail:
            mail.To = 'erik.souza@akersolutions.com; vinicius.oliveira@akersolutions.com ;guilherme.barddal@akersolutions.com'
            mail.Subject = 'Stock Coverage - Script'
            mail.Body = f'Confirmação da execução da rotina de extração e limpeza da data {date.today()}'
            try:
                mail.Attachments.Add(f'{ROOT_DIR}\\{logger.LOG_NAME}')
            except:
                pass
            mail.Send()
            time.sleep(30)
    except:
        pass


if __name__ == '__main__':
    main()
    #os.system("pause")
