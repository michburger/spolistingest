#!/usr/bin/env python3
import argparse
parser = argparse.ArgumentParser(description='Sharepoint Online Ingest Health Check')
parser.add_argument('--log',
                    default='~/logs/spolistingest.log',
                    help='Path to the log file')

try:
    with open(parser.parse_args().log) as f:
        for line in f:
            pass
        last_line = line

        if last_line.startswith('Error'):
            print("Execution failed with error: " + last_line)
            exit(1)

        if not last_line.startswith('SharePoint List Ingest executed at'):
            print("Invalid result text found.")
            exit(1)

        date_str = last_line.split('SharePoint List Ingest executed at ')[1]
        from datetime import datetime, timedelta
        execution_time = datetime.strptime(date_str, '%a %b %d %H:%M:%S %Z %Y')
        if datetime.now() - execution_time > timedelta(minutes=30):
            print("Execution is older than 30 minutes.")
            exit(1)
        else:
            print("Execution is healthy.")

except Exception as ex:
    print(f"Error reading log file: {ex}")
    exit(1)
