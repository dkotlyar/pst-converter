import argparse
import json
import os.path
import datetime

import pypff


class ObjectEncoder(json.JSONEncoder):
    def default(self, obj):
        if hasattr(obj, 'to_json'):
            return obj.to_json()
        if isinstance(obj, (datetime.date, datetime.datetime)):
            return obj.isoformat()

        return super().default(obj)


def folder_traverse(base):
    """
    The folder_traverse function walks through the base of the folder and scans for sub-folders and messages
    :param base: Base folder to scan for new items within the folder.
    :return: None
    """
    messages = []
    for folder in base.sub_folders:
        if folder.number_of_sub_folders:
            messages.extend(folder_traverse(folder))
        messages.extend(check_for_messages(folder))
    return messages


def check_for_messages(folder):
    """
    The check_for_messages function reads folder messages if present and passes them to the report function
    :param folder: pypff.Folder object
    :return: None
    """
    message_list = []
    for message in folder.sub_messages:
        message_dict = process_message(message)
        message_list.append(message_dict)
    return message_list


def process_message(message):
    """
    The process_message function processes multi-field messages to simplify collection of information
    :param message: pypff.Message object
    :return: A dictionary with message fields (values) and their data (keys)
    """
    return {
        "subject": message.subject,
        "sender": message.sender_name,
        "header": message.transport_headers,
        "body": message.plain_text_body,
        "html_body": message.html_body.decode('utf-8'),
        "creation_time": message.creation_time,
        "submit_time": message.client_submit_time,
        "delivery_time": message.delivery_time,
        "attachment_count": message.number_of_attachments,
    }


def parse_file(file, result_dir='./'):
    opst = pypff.open(file)
    root = opst.get_root_folder()
    messages = folder_traverse(root)
    if messages:
        basename = os.path.basename(file)
        name, ext = os.path.splitext(basename)
        os.makedirs(result_dir, exist_ok=True)
        with open(f'{result_dir}/{name}.json', 'w') as f:
            json.dump(messages, f, cls=ObjectEncoder)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('PST_DIR', help="Directory contains files in PST Microsoft Outlook format.")
    parser.add_argument('OUTPUT_DIR', help="Directory of output for json converted files.")
    args = parser.parse_args()

    if not os.path.exists(args.PST_DIR):
        print(f'Directory {args.PST_DIR} doesn\'t exists')
        exit(1)

    if os.path.exists(args.OUTPUT_DIR):
        if os.path.isfile(args.OUTPUT_DIR):
            print(f'Output {args.OUTPUT_DIR} must be directory, not file')
            exit(1)
    else:
        os.mkdir(args.OUTPUT_DIR)

    if os.path.isfile(args.PST_DIR):
        parse_file(args.PST_DIR, args.OUTPUT_DIR)
    else:
        root = f'{args.PST_DIR}/'
        for path, dir, files in os.walk(root):
            subpath = path[len(root):]
            output_path = f'{args.OUTPUT_DIR}/{subpath}/'
            for file in files:
                parse_file(f'{path}/{file}', output_path)

    print('Done')
