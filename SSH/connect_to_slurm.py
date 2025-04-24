import os

import paramiko
import time
import re

import paramiko
import re
import time
hostname='slurm.bgu.ac.il'
port=22
username='yairbary'
password='yairYAIR0_0'
local_file_path = 'C:/Users/yairb/Downloads/elbit_fs3Q.pdf'

def ssh_connect_and_authenticate(hostname=hostname, port=port, username=username, password=password, local_file_path=local_file_path,question_content='print hello world'):
    # Define the question content
    job_output=""

    # Create the question.txt file locally
    question_file_path = "question.txt"
    with open(question_file_path, "w") as question_file:
        question_file.write(question_content)
    print("question.txt file created locally.")

    # Create an SSH client
    client = paramiko.SSHClient()

    # Automatically add the server's host key (not recommended for production)
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    try:
        # Connect to the host
        client.connect(hostname, port=port, username=username, password=password)
        print(f"Connected to {hostname}")
        client.get_transport().set_keepalive(30)  # Keep connection alive

        # Delete the old files if they exist
        delete_command = 'rm -f /sise/home/yairbary/a.pdf /sise/home/yairbary/question.txt'
        client.exec_command(delete_command)
        print("Old files deleted (if they existed).")

        # Upload the local_file_path (PDF file) and question.txt to the server
        sftp = client.open_sftp()
        remote_file_path = "/sise/home/yairbary/a.pdf"
        remote_question_path = "/sise/home/yairbary/question.txt"

        sftp.put(local_file_path, remote_file_path)
        sftp.put(question_file_path, remote_question_path)
        sftp.close()
        print(f"Files {local_file_path} and {question_file_path} uploaded to the server.")

        # Submit the job
        sbatch_command = 'sbatch example.sbatch'
        stdin, stdout, stderr = client.exec_command(sbatch_command)
        sbatch_response = stdout.read().decode()
        print(f"SBATCH Response: {sbatch_response}")

        # Extract the JobId from the response
        job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
        if job_id_match:
            job_id = job_id_match.group(1)
            print(f"Job ID: {job_id}")
        else:
            print("Failed to extract Job ID.")
            return

        # Wait until the job is not running
        while True:
            squeue_command = f"squeue --me | grep {job_id}"
            stdin, stdout, stderr = client.exec_command(squeue_command)
            squeue_response = stdout.read().decode()

            if len(squeue_response) == 0:
                print(f"Job {job_id} is finished.")
                break
            else:
                print(f"Job {job_id} is still running. Waiting...")
                time.sleep(10)  # Wait 10 seconds before checking again

        # Once the job is finished, retrieve the output file
        output_file_command = f"cat job-{job_id}.out"
        stdin, stdout, stderr = client.exec_command(output_file_command)
        job_output = stdout.read().decode()
        response = job_output.split( "Answer the question concisely based on the context provided. use the structure i provided to answer the question:")[-1]
        arr = response.split('\n')
        response = ''

        # Loop through the lines, starting from the second line (index 1) and append non-blank lines
        for line in arr[1:]:
            if line.strip():  # Only append non-blank lines
                response += line + '\n'  # Append the line with a newline



        #print(f"Job {job_id} output:\n{response}")
        file=f"job-{job_id}.out"
        sbatch_command = 'sbatch sen.sbatch'
        stdin, stdout, stderr = client.exec_command(sbatch_command)
        sbatch_response = stdout.read().decode()
        print(f"SBATCH Response: {sbatch_response}")
        job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
        if job_id_match:
            job_id = job_id_match.group(1)
            print(f"Job ID: {job_id}")
        else:
            print("Failed to extract Job ID.")
            return

        # Wait until the job is not running
        while True:
            squeue_command = f"squeue --me | grep {job_id}"
            stdin, stdout, stderr = client.exec_command(squeue_command)
            squeue_response = stdout.read().decode()

            if len(squeue_response) == 0:
                print(f"Job {job_id} is finished.")
                break
            else:
                print(f"Job {job_id} is still running. Waiting...")
                time.sleep(10)  # Wait 10 seconds before checking again
        # Retrieve the sentiment_results.json file from the remote directory
        remote_json_path = "/sise/home/yairbary/sentiment_results.json"
        base_dir = os.path.dirname(os.path.abspath(""))

        # Construct the full path for the sentiment_results.json file inside Presentation/Files
        local_json_path  = os.path.join(base_dir, "Presentation", "Files", "sentiment_results.json")

        sftp = client.open_sftp()
        try:
            sftp.get(remote_json_path, local_json_path)
            print(f"Retrieved {remote_json_path} to {local_json_path}")
        except Exception as e:
            print(f"Failed to retrieve sentiment_results.json: {e}")
        finally:
            sftp.close()

        # Read and parse the JSON file
        import json
        try:
            with open(local_json_path, "r", encoding="utf-8") as json_file:
                sentiment_data = json.load(json_file)
                print(f"Sentiment Results: {json.dumps(sentiment_data, indent=4, ensure_ascii=False)}")
        except Exception as e:
            print(f"Failed to read sentiment_results.json: {e}")

        sbatch_command = 'sbatch summarize.sbatch'
        stdin, stdout, stderr = client.exec_command(sbatch_command)
        sbatch_response = stdout.read().decode()
        print(f"SBATCH Response for summarize job: {sbatch_response}")

        # Extract the JobId from the response for summarize job
        job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
        if job_id_match:
            job_id = job_id_match.group(1)
            print(f"Summarize Job ID: {job_id}")
        else:
            print("Failed to extract Job ID for summarize job.")
            return

        # Wait until the summarize job is not running
        while True:
            squeue_command = f"squeue --me | grep {job_id}"
            stdin, stdout, stderr = client.exec_command(squeue_command)
            squeue_response = stdout.read().decode()

            if len(squeue_response) == 0:
                print(f"Summarize Job {job_id} is finished.")
                break
            else:
                print(f"Summarize Job {job_id} is still running. Waiting...")
                time.sleep(10)  # Wait 10 seconds before checking again

        # Retrieve the company_summary.txt file
        remote_summary_path = "/sise/home/yairbary/company_summary.txt"
        base_dir = os.path.dirname(os.path.abspath(""))

        # Construct the full path for saving the file locally
        local_summary_path = os.path.join(base_dir, "Presentation", "Files", "company_summary.txt")
        sftp = client.open_sftp()
        try:
            sftp.get(remote_summary_path, local_summary_path)
            print(f"Retrieved {remote_summary_path} to {local_summary_path}")
        except Exception as e:
            print(f"Failed to retrieve company_summary.txt: {e}")
        finally:
            sftp.close()
       # client.exec_command(delete_command)
        job_output=response
    except paramiko.AuthenticationException:
        print("Authentication failed, please verify your credentials.")
    except paramiko.SSHException as sshException:
        print(f"Unable to establish SSH connection: {sshException}")
    except Exception as e:
        print(f"Exception in connecting: {e}")
    finally:
        client.close()
    return job_output


#ssh_connect_and_authenticate(hostname, port, username, password, local_file_path,"question_content")