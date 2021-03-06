from RedmineAPI.Utilities import FileExtension, create_time_log
from RedmineAPI.Access import RedmineAccess
from RedmineAPI.Configuration import Setup
import roga_methods
import subprocess
import os

from Utilities import CustomKeys, CustomValues


class Automate(object):

    def __init__(self, force):

        # create a log, can be written to as the process continues
        self.timelog = create_time_log(FileExtension.runner_log)

        # Key: used to index the value to the config file for setup
        # Value: 3 Item Tuple ("default value", ask user" - i.e. True/False, "type of value" - i.e. str, int....)
        # A value of None is the default for all parts except for "Ask" which is True
        # custom_terms = {CustomKeys.key_name: (CustomValues.value_name, True, str)}  # *** can be more than 1 ****
        custom_terms = dict()

        # Create a RedmineAPI setup object to create/read/write to the config file and get default arguments
        setup = Setup(time_log=self.timelog, custom_terms=custom_terms)
        setup.set_api_key(force)

        # Custom terms saved to the config after getting user input
        # self.custom_values = setup.get_custom_term_values()
        # self.your_custom_value_name = self.custom_values[CustomKeys.key_name]

        # Default terms saved to the config after getting user input
        self.seconds_between_checks = setup.seconds_between_check
        self.nas_mnt = setup.nas_mnt
        self.redmine_api_key = setup.api_key

        # Initialize Redmine wrapper
        self.access_redmine = RedmineAccess(self.timelog, self.redmine_api_key)

        self.botmsg = '\n\n_I am a bot. This action was performed automatically._'  # sets bot message
        # Subject name and Status to be searched on Redmine
        self.issue_title = 'autoroga'
        self.issue_status = 'New'

    def timed_retrieve(self):
        """
        Continuously search Redmine in intervals for the inputted period of time, 
        Log errors to the log file as they occur
        """
        import time
        while True:
            # Get issues matching the issue status and subject
            found_issues = self.access_redmine.retrieve_issues(self.issue_status, self.issue_title)
            # Respond to the issues in the list 1 at a time
            while len(found_issues) > 0:
                self.respond_to_issue(found_issues.pop(len(found_issues) - 1))
            self.timelog.time_print("Waiting for the next check.")
            time.sleep(self.seconds_between_checks)

    def respond_to_issue(self, issue):
        """
        Run the desired automation process on the inputted issue, if there is an error update the author
        :param issue: Specified Redmine issue information
        """
        self.timelog.time_print("Found a request to run. Subject: %s. ID: %s" % (issue.subject, str(issue.id)))
        self.timelog.time_print("Adding to the list of responded to requests.")
        self.access_redmine.log_new_issue(issue)

        try:
            issue.redmine_msg = "Beginning run of - %s" % issue.subject
            self.access_redmine.update_status_inprogress(issue, self.botmsg)
            ##########################################################################################
            print("Run your process right here")
            try:
                os.remove('ROGA_summary_OLF.xlsx')
            except FileNotFoundError:
                pass
            try:
                os.remove('ROGA_summary_OLC.xlsx')
            except FileNotFoundError:
                pass
            try:
                os.remove('out.docx')
            except FileNotFoundError:
                pass
            excel_file = self.access_redmine.get_attached_files(issue)
            roga_type, seq_ids= roga_methods.parse_description(issue)
            roga_methods.seq_ids_to_textfile(seq_ids)
            # Now try to download the excel file.
            try:
                address = excel_file[0]['content_url']
                if "ROGA_summary_OLF.xlsx" in address:
                    # Download and write to file.
                    x = self.access_redmine.redmine_api.download_file(address, decode=False)
                    f = open("ROGA_summary_OLF.xlsx", "wb")
                    f.write(x)
                    f.close()
                elif "ROGA_summary_OLC.xlsx" in address:
                    # Download and write to file.
                    x = self.access_redmine.redmine_api.download_file(address, decode=False)
                    f = open("ROGA_summary_OLC.xlsx", "wb")
                    f.write(x)
                else:
                    self.access_redmine.update_issue_to_author(issue, message="ROGA summary excel file not uploaded."
                                                                              " Please create a new issue and try again.")
            except IndexError:
                # If the author didn't upload any files, tell them they messed up.
                self.access_redmine.update_issue_to_author(issue, message="You did not upload any files. Please create a new"
                                                                   " issue, upload a ROGA summary excel file, and "
                                                                   "try again.")
            if roga_type == "listeria":
                f = open("error.txt", "w")
                s = subprocess.run(
                    ["python", "Automate_Listeria.py", "seqIDlist.txt", "out.docx", "-n", self.nas_mnt],
                    stderr=f)
                code = s.returncode
                f.close()
            elif roga_type == "salmonella":
                f = open("error.txt", "w")
                s = subprocess.run(
                    ["python", "Automate_Salmonella.py", "seqIDlist.txt", "out.docx", "-n", self.nas_mnt],
                    stderr=f)
                code = s.returncode
                f.close()
            elif roga_type == "vtec":
                f = open("error.txt", "w")
                s = subprocess.run(
                    ["python", "Automate_VTEC.py", "seqIDlist.txt", "out.docx", "-n", self.nas_mnt], stderr=f)
                code = s.returncode
                f.close()
            elif roga_type == "salmonella_feed":
                f = open("error.txt", "w")
                s = subprocess.run(
                    ["python", "Automate_Salmonella_Feed.py", "seqIDlist.txt", "out.docx", "-n", self.nas_mnt], stderr=f)
                code = s.returncode
                f.close()

            try:
                self.access_redmine.redmine_api.upload_file('out.docx', issue.id,
                                                        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                                        file_name_once_uploaded='ROGA_Report.docx')
                os.remove('out.docx')
            except FileNotFoundError:
                self.access_redmine.update_issue_to_author(issue, message='Report creation was unsuccesful. Uploading'
                                                                          ' error traceback.')
                self.access_redmine.redmine_api.upload_file('error.txt', issue.id,
                                                            'text/plain',
                                                            file_name_once_uploaded='error.txt')
            try:
                os.remove('ROGA_summary_OLF.xlsx')
            except FileNotFoundError:
                pass
            try:
                os.remove('ROGA_summary_OLC.xlsx')
            except FileNotFoundError:
                pass

            ##########################################################################################
            self.completed_response(issue)

        except Exception as e:
            import traceback
            self.timelog.time_print("[Warning] run.py had a problem, continuing redmine api anyways.")
            self.timelog.time_print("[Automation Error Dump]\n" + traceback.format_exc())
            # Send response
            issue.redmine_msg = "There was a problem with your request. Please create a new issue on" \
                                " Redmine to re-run it.\n%s" % traceback.format_exc()
            # Set it to feedback and assign it back to the author
            self.access_redmine.update_issue_to_author(issue, self.botmsg)

    def completed_response(self, issue):
        """
        Update the issue back to the author once the process has finished
        :param issue: Specified Redmine issue the process has been completed on
        :param missing_files: All files that were not correctly uploaded
        """
        # Assign issue back to the author
        self.timelog.time_print("The request to run: %s has been completed." % issue.subject)
        self.timelog.time_print("Assigning the issue back to the author.")

        issue.redmine_msg = "ROGA request complete."

        self.access_redmine.update_issue_to_author(issue, self.botmsg)
        self.timelog.time_print("Completed Response to issue %s." % str(issue.id))
        self.timelog.time_print("The next request will be processed once available")
