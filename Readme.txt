This is a quick tool to automation update the firmware of CopTrax II box. 

Preparation:
1. Download the automation tool from the server and save the whole foder FirmwareAutomation in a thumb drive.
2. Power on the target CopTrax II box. Make sure all the recorded videos have been uploaded before doing the firmware update.
3. If you operate the box with your laptop PC through the remote screen, plug the thumb drive into one of the box's USB port, and then keep typing the Alt-PageUp key until you see the main window. 
4. If you operate the box through 7" USB screen, unplug the front USB camera from the USB port of the box, replug a USB hub with keyboard, mouse, and the USB thumb.

Operation:
1. Copy the whole folder of FirmwareAutomation to the CopTrax II box under c:\;
2. Go to the folder C:\FirmwareAutomation;
3. Double click FirmwareUpdateTool.exe;
4. The tool will run to update the firmware automatically. The progresses will be shown up on pop-up windows. There going to be several auto-reboot during the updating. 
5. When see the message that requires for hard reset or the LED on the box turn steadly red, press the hard reset button.
6. After the firmware been updated successfully, the Coptrax II application will be started. All the trails including the folder C:\FirmwareAutomation will be cleanup.
7. When some fatal failure occurs, the automation will stop. You may try to run the automation from step 1 again. If there is still fatal errors, you have to update the firmware manually.
8. To update the firmware to another version, you can modify the version.cfg file with a text editor, like notpad.exe, to replace the version with new targetversion like x.x.x. Make sure the copy the corresponding target hex file in C:\FirmwareAutomation. The hex filename shall be in format "TriggerBox 2.0 App x.x.x.hex.
