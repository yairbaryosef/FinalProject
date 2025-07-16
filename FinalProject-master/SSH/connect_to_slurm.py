import os
import paramiko
import re
import time
import json
def run_it_all(name, base_dir):
    print('test address', os.path.join(base_dir, "Presentation", "Files", "sentiment_results.json"))
    hostname = 'slurm.bgu.ac.il'
    port = 22
    username = 'yairbary'
    password = 'yairYAIR0_00'
    local_file_path = name

    def ssh_connect_and_authenticate(
        hostname=hostname,
        port=port,
        username=username,
        password=password,
        local_file_path=local_file_path,
        question_content='print hello world'
    ):
        job_output = ""
        question_file_path = "question.txt"

        with open(question_file_path, "w") as question_file:
            question_file.write(question_content)
        print("question.txt file created locally.")

        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        try:
            client.connect(hostname, port=port, username=username, password=password)
            print(f"Connected to {hostname}")
            client.get_transport().set_keepalive(30)

            delete_command = 'rm -f /sise/home/yairbary/a.pdf /sise/home/yairbary/question.txt'
            client.exec_command(delete_command)

            sftp = client.open_sftp()
            remote_file_path = "/sise/home/yairbary/a.pdf"
            remote_question_path = "/sise/home/yairbary/question.txt"
            sftp.put(local_file_path, remote_file_path)
            sftp.put(question_file_path, remote_question_path)
            sftp.close()
            print("Files uploaded to the server.")

            # מיקום הפלטים (קשיח)
            #local_output_dir = r"C:\Users\orel yosef\Desktop\מלגו\FinalProject-master\SSH"


            # --- Submit example job ---
            sbatch_command = 'sbatch example.sbatch'
            stdin, stdout, stderr = client.exec_command(sbatch_command)
            sbatch_response = stdout.read().decode()
            job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
            if not job_id_match:
                print("Failed to extract Job ID.")
                return
            job_id = job_id_match.group(1)

            while True:
                stdin, stdout, stderr = client.exec_command(f"squeue --me | grep {job_id}")
                if stdout.read().decode() == "":
                    break
                time.sleep(10)

            stdin, stdout, stderr = client.exec_command(f"cat job-{job_id}.out")
            job_output = stdout.read().decode()
            response = job_output.split("Answer the question concisely based on the context provided. use the structure i provided to answer the question:")[-1]
            arr = response.split('\n')
            response = ''.join([line + '\n' for line in arr[1:] if line.strip()])

            # --- Sentiment job ---
            sbatch_command = 'sbatch sen.sbatch'
            stdin, stdout, stderr = client.exec_command(sbatch_command)
            sbatch_response = stdout.read().decode()
            job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
            if not job_id_match:
                return
            job_id = job_id_match.group(1)
    
            while True:
                stdin, stdout, stderr = client.exec_command(f"squeue --me | grep {job_id}")
                if stdout.read().decode() == "":
                    break
                time.sleep(10)
    
            sftp = client.open_sftp()
            remote_json_path = "/sise/home/yairbary/sentiment_results.json"
            local_json_path = os.path.join(base_dir, "Presentation", "Files", "sentiment_results.json")
            try:
                sftp.get(remote_json_path, local_json_path)
                print(f"Saved sentiment_results.json to {local_json_path}")
            except Exception as e:
                print(f"Error: {e}")
            sftp.close()

            # --- LLM_LINES job ---
            local_llm_output = os.path.join(base_dir, "Presentation", "Files", "llm_lines.txt")
            print(local_llm_output)
            sbatch_command = 'sbatch llm_lines.sbatch'
            stdin, stdout, stderr = client.exec_command(sbatch_command)
            sbatch_response = stdout.read().decode()
            job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
            if not job_id_match:
                print("Failed to submit llm_lines job.")
            else:
                job_id = job_id_match.group(1)
                print(f"Submitted llm_lines job with ID {job_id}")

                # המתנה לסיום הריצה
                while True:
                    stdin, stdout, stderr = client.exec_command(f"squeue --me | grep {job_id}")
                    if stdout.read().decode() == "":
                        break
                    time.sleep(10)

                # הורדת הפלט של llm_lines.txt
                sftp = client.open_sftp()
                remote_llm_output = f"/sise/home/yairbary/llm_lines.txt"

                try:
                    sftp.get(remote_llm_output, local_llm_output)
                    print(f"✅ Saved LLM output to {local_llm_output}")
                except Exception as e:
                    print(f"❌ Failed to download llm_lines.txt: {e}")
                sftp.close()



            # --- SWOT job ---
            '''sbatch_command = 'sbatch swot.sbatch'
            stdin, stdout, stderr = client.exec_command(sbatch_command)
            sbatch_response = stdout.read().decode()
            job_id_match = re.search(r'Submitted batch job (\d+)', sbatch_response)
            if not job_id_match:
                return
            job_id = job_id_match.group(1)
    
            while True:
                stdin, stdout, stderr = client.exec_command(f"squeue --me | grep {job_id}")
                if stdout.read().decode() == "":
                    break
                time.sleep(10)
    
            sftp = client.open_sftp()
            remote_swot_output_path = f"/sise/home/yairbary/swot_output.txt"
            local_summary_path = os.path.join(base_dir, "Presentation", "Files", "swot_output.txt")
    
            try:
                sftp.get(remote_swot_output_path, local_summary_path)
                print(f"Saved SWOT output to {local_summary_path}")
            except Exception as e:
                print(f"Error: {e}")
            sftp.close()
            '''
            job_output = ''

        except Exception as e:
            print(f"Exception: {e}")


        finally:
            client.close()

        return job_output

    # להפעיל:
    ssh_connect_and_authenticate()
