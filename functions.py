from datetime import datetime, timedelta
import pandas as pd
import win32com.client
import glob


def make_df(read_directory: str) -> pd.DataFrame:
    """
    This function uses several other functions to read sort files from disk, then creates
    a dataframe. Will write csv to disk if output_csv is True. 
    
    Args:
        read_directory (str): Directory to read sort files from.
    Returns:
        df (pd.DataFrame): Pandas dataframe of observations from disk. 
    """
    sort_files_dict = sort_files(read_directory)

    passdown_list_of_lists = []

    for status_file, gauge_file in sort_files_dict.items():
        try:
            sort_instance = passdown_generator(status_file, gauge_file)
            passdown_list_of_lists.append(sort_instance)
        except FileNotFoundError:
            continue

    cols = ['timestamp', 'sort_id', 'sort', 'weekday', 'volume', 'op_reject', 'iss_reject', 'scan_tunnel_reject',
            'mechanical_reject', 'ss1_lane_full', 'ss2_lane_full', 'lane_full', 'ss1_iss_not_on_file',
            'ss2_iss_not_on_file', 'iss_not_on_file', 'ss1_iss_alt_not_on_file', 'ss2_iss_alt_not_on_file',
            'iss_alt_not_on_file', 'ss1_iss_unassigned_dest', 'ss2_iss_unassigned_dest', 'iss_unassigned_dest',
            'ss1_iss_unassigned_nlpt', 'ss2_iss_unassigned_nlpt', 'iss_unassigned_nlpt', 'ss1_iss_unassigned_trailer',
            'ss2_iss_unassigned_trailer', 'iss_unassigned_trailer', 'ss1_iss_no_response', 'ss2_iss_no_response',
            'iss_no_response', 'ss1_iss_late_response', 'ss2_iss_late_response', 'iss_late_response',
            'ss1_invalid_asgn_to_plc', 'ss2_invalid_asgn_to_plc', 'invalid_asgn_to_plc', 'ss1_invalid_destination',
            'ss2_invalid_destination', 'invalid_destination', 'ss1_no_read', 'ss2_no_read', 'no_read', 'ss1_multi_read',
            'ss2_multi_read', 'multi_read', 'ss1_bad_xmit', 'ss2_bad_xmit', 'bad_xmit', 'ss1_no_xmit', 'ss2_no_xmit',
            'no_xmit', 'ss1_chute_jam', 'ss2_chute_jam', 'chute_jam', 'ss1_chute_chute_disabled', 'ss2_chute_disabled',
            'ss1_diverter_fault', 'ss2_diverter_fault', 'diverter_fault', 'ss1_divert_failed', 'ss2_diver_failed',
            'divert_failed', 'ss1_divert_inhibit', 'ss2_divert_inhibit', 'divert_inhibit', 'ss1_gap_error',
            'ss2_gap_error', 'gap_error', 'ss1_lost_tracking', 'ss2_lost_tracking', 'lost_tracking',
            'ss1_sorter_not_at_speed', 'ss2_sorter_not_at_speed', 'sorter_not_at_speed', 'ss1_secondary_no_show',
            'ss2_secondary_no_show', 'secondary_no_show', 'ss1_divert_out_of_position', 'ss2_divert_out_of_position',
            'divert_out_of_position', 'ss1_sorter_aux_mode', 'ss2_sorter_aux_mode', 'sorter_aux_mode']

    df = pd.DataFrame(passdown_list_of_lists, columns=cols)

    df = df.set_index('timestamp').sort_index()
    return df


def sort_files(directory: str) -> dict:
    """
    This function will return a list of all sort_status files that have a matching sort_gauge counterpart.
    
    Args:
        directory (str): Navigation directory for file reading.

    Returns:
        id_matches (list): List of all sort_status files for which there are matching pairs.
    """
    # List all files in directory using pattern matching.
    files = glob.glob(directory)

    # Creating a set of Sorter Status files from files.
    status_tuple = (status_file for status_file in files if 'Status' in status_file)

    # Initializing dictionary to store sorter_status and sort_gauge file pairs.
    match_dict = {}

    # Creating key value pairs, key is sorter_status value is sort_gauge.
    for status_file in status_tuple:
        match_dict[status_file] = status_file.replace('erStatus', 'Gauge')

    return match_dict


def sort_name(hour: int) -> str:
    sort = None
    if hour <= 3:
        sort = 'twi'
    elif hour <= 14:
        sort = 'pre'
    elif hour <= 17:
        sort = 'day'
    elif hour <= 24:
        sort = 'twi'
    return sort


def extractor(start_point, read_bytes, file_select):
    file_select.seek(start_point)
    attribute = file_select.read(read_bytes)
    return attribute


def passdown_generator(status_file: str, gauge_file: str) -> list:
    """
    Reads sorter_status and sort_gauge files and collects certain datapoints to create an observation. 
    Args:
        status_file (str): This is the name of the Sorter Status file.
        gauge_file (str): This is the name of the Sort Gauge file.
    Returns:
        return_list (list): This is the observations accrued from opening the sorter_status and sort_gauge
        files. 
    """
    sort_id = status_file[39:46]
    sorter_status = open(f"{status_file}", "r")
    sort_gauge = open(f"{gauge_file}", "r")

    date = datetime.strptime(extractor(0, 10, sorter_status), '%m/%d/%Y')
    time = datetime.strptime(extractor(80, 8, sorter_status), "%H:%M:%S").time()
    # Compensating for late dock closures on twi sorts.
    if time.hour <= 3:
        date = date - timedelta(days=1)
    timestamp = datetime.combine(date, time)
    weekday = date.strftime("%A")

    sort = sort_name(timestamp.hour)

    # Deprecated.
    # ss1_running_time = extractor(4363, 5)
    # ss2_running_time = extractor(4369, 5)
    # runtime_minutes = (int(ss1_running_time[0:2]) + int(ss2_running_time[0:2])) * 60 + \
    #                  int(ss1_running_time[3:]) + int(ss2_running_time[3:])
    # smalls_volume = extractor(514, 5, sort_gauge)
    # secondary_inducts = int(extractor(320, 10))
    # rehandle = int(extractor(4967, 6))

    # Sort Statistics
    volume = int(extractor(288, 6, sort_gauge))

    # Operational Reject = Lane Full, Load Oversize, Wrong Sorter, Disable Sortation
    ss1_lane_full = int(extractor(5898, 5, sorter_status))
    ss2_lane_full = int(extractor(5905, 5, sorter_status))
    lane_full = ss1_lane_full + ss2_lane_full

    # ISS Reject = Barcode Not on File, Alt Barcode Not on File, Unassign Destination, 
    # Unassign Next Load Point, Unassign Trailer not open, ISS No Response, 
    # ISS Late Response, Invalid Assignment to PLC, Invalid Destination Terminal
    ss1_iss_not_on_file = int(extractor(4999, 4, sorter_status))
    ss2_iss_not_on_file = int(extractor(5005, 4, sorter_status))
    iss_not_on_file = ss1_iss_not_on_file + ss1_iss_not_on_file

    ss1_iss_alt_not_on_file = int(extractor(5073, 5, sorter_status))
    ss2_iss_alt_not_on_file = int(extractor(5079, 5, sorter_status))
    iss_alt_not_on_file = ss1_iss_alt_not_on_file + ss2_iss_alt_not_on_file

    ss1_iss_unassigned_dest = int(extractor(5148, 5, sorter_status))
    ss2_iss_unassigned_dest = int(extractor(5154, 5, sorter_status))
    iss_unassigned_dest = ss1_iss_unassigned_dest + ss2_iss_unassigned_dest

    ss1_iss_unassigned_nlpt = (int(extractor(5224, 5, sorter_status)))
    ss2_iss_unassigned_nlpt = (int(extractor(5229, 5, sorter_status)))
    iss_unassigned_nlpt = ss1_iss_unassigned_nlpt + ss2_iss_unassigned_nlpt

    ss1_iss_unassigned_trailer = int(extractor(5298, 5, sorter_status))
    ss2_iss_unassigned_trailer = int(extractor(5304, 5, sorter_status))
    iss_unassigned_trailer = ss1_iss_unassigned_trailer + ss2_iss_unassigned_trailer

    ss1_iss_no_response = int(extractor(5749, 5, sorter_status))
    ss2_iss_no_response = int(extractor(5754, 5, sorter_status))
    iss_no_response = ss1_iss_no_response + ss2_iss_no_response

    ss1_iss_late_response = int(extractor(5824, 5, sorter_status))
    ss2_iss_late_response = int(extractor(5829, 5, sorter_status))
    iss_late_response = ss1_iss_late_response + ss2_iss_late_response

    ss1_invalid_asgn_to_plc = int(extractor(5673, 5, sorter_status))
    ss2_invalid_asgn_to_plc = int(extractor(5679, 5, sorter_status))
    invalid_asgn_to_plc = ss1_invalid_asgn_to_plc + ss2_invalid_asgn_to_plc

    ss1_invalid_destination = int(extractor(7023, 5, sorter_status))
    ss2_invalid_destination = int(extractor(7029, 5, sorter_status))
    invalid_destination = ss1_invalid_destination + ss2_invalid_destination

    # Scan Tunnel Reject = No Read, Multi Read, Scan Tunnel Bad Transmit, Scan Tunnel No Transmit
    ss1_no_read = int(extractor(5373, 6, sorter_status))
    ss2_no_read = int(extractor(5379, 7, sorter_status))
    no_read = ss1_no_read + ss2_no_read

    ss1_multi_read = int(extractor(5448, 5, sorter_status))
    ss2_multi_read = int(extractor(5454, 5, sorter_status))
    multi_read = ss1_multi_read + ss2_multi_read

    ss1_bad_xmit = int(extractor(7098, 5, sorter_status))
    ss2_bad_xmit = int(extractor(7104, 5, sorter_status))
    bad_xmit = ss1_bad_xmit + ss2_bad_xmit

    ss1_no_xmit = int(extractor(7174, 6, sorter_status))
    ss2_no_xmit = int(extractor(7180, 6, sorter_status))
    no_xmit = ss1_no_xmit + ss2_no_xmit

    # Mechanical Reject = Chute Jam, Chute Disabled, Diverter Fault, Diverter Failed, 
    # Divert Inhibit due to FMS, Gap Error, Lost Tracking, Sorter Not at Speed, 
    # Secondary No-show, Update PE Length Change, Divert Out of Position, Aux Sorter Unavailable
    ss1_chute_jam = int(extractor(5973, 5, sorter_status))
    ss2_chute_jam = int(extractor(5979, 5, sorter_status))
    chute_jam = ss1_chute_jam + ss2_chute_jam

    ss1_chute_disabled = int(extractor(6049, 5, sorter_status))
    ss2_chute_disabled = int(extractor(6055, 5, sorter_status))
    chute_disabled = ss1_chute_disabled + ss2_chute_disabled

    ss1_diverter_fault = int(extractor(6123, 5, sorter_status))
    ss2_diverter_fault = int(extractor(6129, 5, sorter_status))
    diverter_fault = ss1_diverter_fault + ss2_diverter_fault

    ss1_divert_failed = int(extractor(6198, 5, sorter_status))
    ss2_divert_failed = int(extractor(6204, 5, sorter_status))
    divert_failed = ss1_divert_failed + ss2_divert_failed

    ss1_divert_inhibit = int(extractor(6273, 5, sorter_status))
    ss2_divert_inhibit = int(extractor(6279, 5, sorter_status))
    divert_inhibit = ss1_divert_inhibit + ss2_divert_inhibit

    ss1_gap_error = int(extractor(6348, 5, sorter_status))
    ss2_gap_error = int(extractor(6354, 5, sorter_status))
    gap_error = ss1_gap_error + ss2_gap_error

    ss1_lost_tracking = int(extractor(6424, 5, sorter_status))
    ss2_lost_tracking = int(extractor(6430, 5, sorter_status))
    lost_tracking = ss1_lost_tracking + ss2_lost_tracking

    ss1_sorter_not_at_speed = int(extractor(6498, 5, sorter_status))
    ss2_sorter_not_at_speed = int(extractor(6504, 5, sorter_status))
    sorter_not_at_speed = ss1_sorter_not_at_speed + ss2_sorter_not_at_speed

    ss1_secondary_no_show = int(extractor(5598, 5, sorter_status))
    ss2_secondary_no_show = int(extractor(5604, 5, sorter_status))
    secondary_no_show = ss1_secondary_no_show + ss2_secondary_no_show

    ss1_divert_out_position = int(extractor(7473, 5, sorter_status))
    ss2_divert_out_position = int(extractor(7479, 5, sorter_status))
    divert_out_position = ss1_divert_out_position + ss2_divert_out_position

    ss1_sorter_aux_mode = int(extractor(7546, 4, sorter_status))
    ss2_sorter_aux_mode = int(extractor(7552, 4, sorter_status))
    sorter_aux_mode = ss1_sorter_aux_mode + ss2_sorter_aux_mode

    # PLC Reject = Package not verified, PLC Missort, Intentional No-Sort 

    sort_gauge.close()
    sorter_status.close()

    # Collecting the various types of rejects.
    op_reject = lane_full

    iss_reject = iss_not_on_file + iss_alt_not_on_file + iss_unassigned_dest + iss_unassigned_nlpt + \
        iss_unassigned_trailer + iss_no_response + iss_late_response + invalid_asgn_to_plc + \
        invalid_destination

    scan_tunnel_reject = no_read + multi_read + bad_xmit + no_xmit

    mechanical_reject = chute_jam + chute_disabled + diverter_fault + divert_failed + divert_inhibit + \
        gap_error + lost_tracking + sorter_not_at_speed + secondary_no_show + divert_out_position + sorter_aux_mode

    columns = [timestamp, sort_id, sort, weekday, volume, op_reject, iss_reject, scan_tunnel_reject, mechanical_reject,
               ss1_lane_full, ss2_lane_full, lane_full,
               ss1_iss_not_on_file, ss2_iss_not_on_file, iss_not_on_file,
               ss1_iss_alt_not_on_file, ss2_iss_alt_not_on_file, iss_alt_not_on_file,
               ss1_iss_unassigned_dest, ss2_iss_unassigned_dest, iss_unassigned_dest,
               ss1_iss_unassigned_nlpt, ss2_iss_unassigned_nlpt, iss_unassigned_nlpt,
               ss1_iss_unassigned_trailer, ss2_iss_unassigned_trailer, iss_unassigned_trailer,
               ss1_iss_no_response, ss2_iss_no_response, iss_no_response,
               ss1_iss_late_response, ss2_iss_late_response, iss_late_response,
               ss1_invalid_asgn_to_plc, ss2_invalid_asgn_to_plc, invalid_asgn_to_plc,
               ss1_invalid_destination, ss2_invalid_destination, invalid_destination,
               ss1_no_read, ss2_no_read, no_read,
               ss1_multi_read, ss2_multi_read, multi_read,
               ss1_bad_xmit, ss2_bad_xmit, bad_xmit,
               ss1_no_xmit, ss2_no_xmit, no_xmit,
               ss1_chute_jam, ss2_chute_jam, chute_jam,
               ss1_chute_disabled, ss2_chute_disabled,
               ss1_diverter_fault, ss2_diverter_fault, diverter_fault,
               ss1_divert_failed, ss2_divert_failed, divert_failed,
               ss1_divert_inhibit, ss2_divert_inhibit, divert_inhibit,
               ss1_gap_error, ss2_gap_error, gap_error,
               ss1_lost_tracking, ss2_lost_tracking, lost_tracking,
               ss1_sorter_not_at_speed, ss2_sorter_not_at_speed, sorter_not_at_speed,
               ss1_secondary_no_show, ss2_secondary_no_show, secondary_no_show,
               ss1_divert_out_position, ss2_divert_out_position, divert_out_position,
               ss1_sorter_aux_mode, ss2_sorter_aux_mode, sorter_aux_mode]

    return columns


def outlook_attachments() -> None:
    files_list = glob.glob('sort_files/*.txt')
    files_list = [x[11:] for x in files_list]
    # Connect to outlook.
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # Navigate to the Oasis folder.
    oasis_folder = outlook.GetDefaultFolder(4).folders("Oasis")
    # Define items inside Oasis folder.
    messages = oasis_folder.Items
    # Start reading messages in Oasis folder, start from 0.
    message = messages.GetFirst()
    # Defining path to save attachment to.
    path = 'fxg-notebooks/sort_files/'
    # Iterating through all Oasis messages, saving files that have not yet been saved.
    new_file_count = 0
    message_count = 0

    while True:
        # print('Messages processed: ', message_count, end='\r')
        message_count += 1
        # Read attachment from outlook message.
        try:
            attachment = message.Attachments.Item(1)
        # If GetNext() returned None, all messages have been read, break loop.
        except AttributeError:
            break
        # Check for attachment membership in files_list.
        try:
            if files_list.index(str(attachment)) >= 0:
                message = messages.GetNext()
                continue
        # If no membership, save attachment.
        except ValueError:
            attachment.SaveASFile(path + str(attachment))
            message = messages.GetNext()
            new_file_count += 1

    return None
