import sys
import pandas as pd


def read_csv_data(file):
    '''Returns a dataframe with variables for each email'''
    with open('_read_files/' + file) as f:
        ws = pd.read_csv(f)
    return ws


def read_email(file):
    '''Returns content of template email'''
    with open('_read_files/' + file, 'r') as email_file:
        email_content = email_file.read()
    return email_content


def write_emails(df, email_content):
    '''Formats emails with rows from dataframe'''
    with open('_emails/output.txt', 'w') as new_file:
        for index, row in df.iterrows():
            formatted_message = email_content.format(
                row['var1'],
                row['var2'],
                row['var3'])
            new_file.write(formatted_message+'\n'+('='*50)+'\n'+'\n')
    return new_file


def main():
    csv = 'example_email_list.csv'
    email_file = 'email.txt'
    df = read_csv_data(csv)
    content = read_email(email_file)
    write_emails(df, content)


if __name__ == '__main__':
    main()
