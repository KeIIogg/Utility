#!/usr/bin/bash
pkg update;
termux-setup-storage;
pkg install proot-distro pulseaudio vim;

# Add line for exucute pulse audio at start up

 sed -i 's/$/pulseaudio --start --load="module-native-protocol-tcp auth-ip-acl=127.0.0.1 auth-anonymous=1" --exit-idle-time=-1
pacmd load-module module-native-protocol-tcp auth-ip-acl=127.0.0.1 auth-anonymous=1/g' ~/.profile;

proot-distro install debian;
proot-distro login debian --user root --shared-tmp;
apt update;
apt install sudo vim firefox-esr;
apt install kde-full;
ln -sf /usr/share/zoneinfo/Asia/Seoul /etc/localtime;
apt install locales;

sed -i 's/#en_US.UTF-8 UTF-8/en_US.UTF-8 UTF-8/g' /etc/locale.gen;
locale-gen;
echo "LANG=en_US.UTF-8" > /etc/locale.conf
passwd;
groupadd storage;
groupadd wheel;
groupadd video;
useradd -m -g users -G wheel,audio,video,storage -s /bin/bash user;
passwd user;

sed -i 's/$/user ALL=(ALL:ALL) ALL/g' /etc/sudoers;  
su user;
cd ~;
sed -i 'export PULSE_SERVER=127.0.0.1 && pulseaudio --start --disable-shm=1 --exit-idle-time=-1' ~/.profile; 
sed -i 'export DISPLAY=:0

# For XFCE4 desktop
dbus-launch --exit-with-session startxfce4 &

# For KDE Plasma desktop
dbus-launch --exit-with-session startplasma-x11 &' ~/startx.sh
chmod +x startx.sh;
exit;



