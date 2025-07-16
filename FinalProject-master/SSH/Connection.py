import paramiko

def ssh_connect_and_authenticate(hostname, port, username, password):
    # Create an SSH client
    client = paramiko.SSHClient()

    # Automatically add the server's host key (not recommended for production)
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    try:
        # Connect to the host
        client.connect(hostname, port=port, username=username, password=password)
        print(f"Connected to {hostname}")

        # Execute a command and get the output
        stdin, stdout, stderr = client.exec_command('ls')
        print("Server response:", stdout.read().decode())
        client.exec_command('cat /opt/conda-def-4-bashrc >> ~/.bashrc')
        client.exec_command('source ~/.bashrc')
        stdin, stdout, stderr=  client.exec_command('python m.py')
        print("Server response:", stdout.read().decode())

    except paramiko.AuthenticationException:
        print("Authentication failed, please verify your credentials.")
    except paramiko.SSHException as sshException:
        print(f"Unable to establish SSH connection: {sshException}")
    except Exception as e:
        print(f"Exception in connecting: {e}")
    finally:
        client.close()


# Usage

# Usage
hostname = '132.72.64.210'
port = 22  # Default SSH port
username = 'yairbary'
password = 'yairYAIR0_0'

ssh_connect_and_authenticate(hostname, port, username, password)
