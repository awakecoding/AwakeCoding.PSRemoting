# Test SSH Server with PowerShell Remoting
FROM ubuntu:22.04

# Avoid interactive prompts during build
ENV DEBIAN_FRONTEND=noninteractive

# Install OpenSSH Server and PowerShell
RUN apt-get update && apt-get install -y \
    openssh-server \
    wget \
    apt-transport-https \
    software-properties-common \
    && wget -q https://packages.microsoft.com/config/ubuntu/22.04/packages-microsoft-prod.deb \
    && dpkg -i packages-microsoft-prod.deb \
    && apt-get update \
    && apt-get install -y powershell \
    && rm -rf /var/lib/apt/lists/* \
    && rm packages-microsoft-prod.deb

# Create test user with known credentials
RUN useradd -m -s /bin/bash testuser \
    && echo 'testuser:testpass' | chpasswd

# Configure SSH
RUN mkdir -p /var/run/sshd \
    && sed -i 's/#PasswordAuthentication yes/PasswordAuthentication yes/' /etc/ssh/sshd_config \
    && sed -i 's/#PermitRootLogin prohibit-password/PermitRootLogin no/' /etc/ssh/sshd_config \
    && echo 'Subsystem powershell /usr/bin/pwsh -NoLogo -NoProfile -s' >> /etc/ssh/sshd_config \
    && echo 'PubkeyAuthentication yes' >> /etc/ssh/sshd_config

# Generate host keys
RUN ssh-keygen -A

# Create .ssh directory for testuser
RUN mkdir -p /home/testuser/.ssh \
    && chmod 700 /home/testuser/.ssh \
    && chown testuser:testuser /home/testuser/.ssh

EXPOSE 22

CMD ["/usr/sbin/sshd", "-D", "-e"]
