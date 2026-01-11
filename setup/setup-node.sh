#!/bin/bash
# Setup a new Pi node in the BlackRoad cluster

HOSTNAME=$1
IP=$2

if [ -z "$HOSTNAME" ] || [ -z "$IP" ]; then
    echo "Usage: ./setup-node.sh <hostname> <ip>"
    exit 1
fi

echo "Setting up $HOSTNAME at $IP..."

ssh pi@$IP << REMOTE
# Set hostname
sudo hostnamectl set-hostname $HOSTNAME

# Install avahi for mDNS
sudo apt-get update
sudo apt-get install -y avahi-daemon nginx

# Enable services
sudo systemctl enable avahi-daemon nginx
sudo systemctl start avahi-daemon nginx

echo "Node $HOSTNAME ready!"
REMOTE
