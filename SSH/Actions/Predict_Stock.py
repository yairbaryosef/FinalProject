import paramiko
import os
import time
import re

def submit_lstm_job(hostname, port, username, password,model='LSTM'):

    local_stock_file = "stock.txt"
    remote_stock_file = "/sise/home/yairbary/stock.txt"
    remote_output_dir = "/sise/home/yairbary/output"
    files_to_download = ["results.txt", "simulated_forecast.png"]

    # Write 'AMZN' into stock.txt


    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    try:
        # Connect to remote server
        client.connect(hostname, port=port, username=username, password=password)
        print(f"Connected to {hostname}")
        client.get_transport().set_keepalive(30)

        # Delete old stock.txt if exists
        delete_stock_command = f'rm -f {remote_stock_file}'
        client.exec_command(delete_stock_command)
        print("Old stock.txt deleted if it existed.")

        # Upload new stock.txt
        sftp = client.open_sftp()
        sftp.put(local_stock_file, remote_stock_file)
        print("stock.txt uploaded to the server.")
        sftp.close()

        # Submit LSTM job
        if model == "LSTM":
          sbatch_command = 'sbatch LSTM.sbatch'
        else:
            sbatch_command='sbatch LSTM_Fine_Tuning.sbatch'

        stdin, stdout, stderr = client.exec_command(sbatch_command)
        sbatch_response = stdout.read().decode()
        print(f"SBATCH Response: {sbatch_response}")

        # Extract job ID
        job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
        if job_id_match:
            job_id = job_id_match.group(1)
            print(f"Job ID: {job_id}")
        else:
            print("Failed to extract Job ID.")
            return

        # Wait for job to finish
        while True:
            check_command = f"squeue --me | grep {job_id}"
            stdin, stdout, stderr = client.exec_command(check_command)
            job_status = stdout.read().decode()

            if not job_status:
                print(f"Job {job_id} is finished.")
                break
            else:
                print(f"Job {job_id} still running...")
                time.sleep(10)

        # Download the specified output files

        # Download the specified output files
        local_output_dir = os.path.join(os.getcwd(), "output")
        os.makedirs(local_output_dir, exist_ok=True)

        sftp = client.open_sftp()
        try:
            for filename in files_to_download:
                remote_file_path = f"{remote_output_dir}/{filename}"  # safer for remote paths
                local_file_path = os.path.join(local_output_dir, filename)
                sftp.get(remote_file_path, local_file_path)
                print(f"Downloaded {filename} to {local_file_path}")
        except Exception as e:
            print(f"Failed to retrieve output files: {e}")
        finally:
            sftp.close()


    except paramiko.AuthenticationException:
        print("Authentication failed. Please check your credentials.")
    except paramiko.SSHException as sshException:
        print(f"SSH error: {sshException}")
    except Exception as e:
        print(f"Error occurred: {e}")
    finally:
        client.close()
        print("SSH connection closed.")
