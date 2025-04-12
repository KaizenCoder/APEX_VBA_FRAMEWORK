#!/bin/bash
# Script pour configurer wsl.conf

# Créer la configuration
cat << 'EOF' | sudo tee /etc/wsl.conf
[boot]
systemd=true

[automount]
enabled = true
options = "metadata,umask=22,fmask=11"
mountFsTab = false

[interop]
enabled = true
appendWindowsPath = true
EOF

echo "Configuration WSL mise à jour. Redémarrez WSL avec 'wsl --shutdown' pour appliquer les changements." 