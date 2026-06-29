# AA-SI_Echoview
[Echoview](https://echoview.com/) is a commercially-available software package used throughout NOAA to process active acoustic data. In Fisheries, Echoview is commonly used and can be an integral component to each Science Centers' data workflow. Echoview operates solely on the Windows operating system and requires a license to run. The license can be accessed via a hardware key (aka "dongle") or a digital license. For local operation on a PC, most users have a hardware key. This allows one user at a time to use that license - i.e., the hardware key must be physically connected to the PC and must be given to another user to share that license. An advantage to the hardware key is that no internet connection is required to run Echoview. A disadvantage is that only one person can use it at a time and sharing that license requires providing the physical dongle to the other user.

The digital licenses require an internet connection because the license validity is checked at run time to a license manager on the network. Advantages are that the license can be shared among users without physical transfer of the hardware key, and depending on how many "seats" the license has, can be used simultaneously among users. A disadvantage is that a network connection is required, which may not always be available - e.g., when at sea. A digital license can be run on a PC, or in "the cloud". For purposes here, we will consider the cloud license as the digital license.

The AA-SI currently is using a timed cloud license on a version of the PAM Windows VM. This license has three seats, meaning three people can be using Echoview simultaneously. We hope to have an annual license soon.  For more information on the Windows VM, see the [AA-SI GitHub Windows VM repo](https://github.com/nmfs-ost/AA-SI_WindowsVM) </br>

The cloud-licensing instructions are found in the PDF document, "Echoview cloud licensing instructions.pdf" that can be accessed via: "P:\Echoview\Echoview cloud licensing instructions.pdf" (on the "dev" GCS environment) or here [Echoview Cloud License](./docs/Echoview_Echoview_cloud_licensing_instructions.pdf). If you run Echoview and the license is not found or working, please see these instructions. </br>

Echoview essentially works the same on the AA-SI Windows VM as it does on a local PC, with one exception (that we've found so far) - that is the creation of .evi files. See the section [evi files](#evi-files) for more information. 

## Link to GCS prod Bucket
Currently, the default drive mount is to the GCS "dev" bucket when you start the Windows VM. In the future, this will be modified to the "prod" bucket. In the meantime, create a new mount to the "prod" bucket:
1. Follow the instructions in [AA-SI_WindowsVM](https://github.com/nmfs-ost/AA-SI_WindowsVM/tree/main/README.md#Using-the-Windows-VM) to create the mount to the prod bucket.
2. When the AA-SI gets a full-fledged Windows VM, some of this will change (e.g., default to the prod, rather than dev bucket), but for now the instructions work - they are just clunky.

## evi Files
When Echoview first reads a data file, or when the version of Echoview has changed such that the format of the .evi file changes, Echoview creates a .evi file. This file is used by Echoview to read a data file more efficiently and faster than an initial read. These .evi files are by default written to the same directory as the data files. On a local PC this is not a problem. On local networks, this process can slow down initial reads of data when/if there are network interuptions. However, reading data files and writing the .evi files from the Windows VM causes severe limitations. For example, in initial testing we found that it took 25 minutes to read one 200-MB EK60 .raw file! After the .evi file is written, data reading is as fast as on local PCs or networks for subsequent reads. We are working with Echoview to modify how and where they write the .evi files, but in the interim, Daniel Woodrich created an alternative method that works well, but requires a few more steps. </br>

1. The instructions are provided in this [idx fix document](./docs/Echoview_idx_files_issue_immediate_solution.pdf).
1. A version of the rclone_custom.conf file for the prod GCS bucket is [here](docs/rclone_custom_prod.conf.txt).
   1. Make sure you remove the ".txt" suffix!
1. Open a command prompt window by either clicking on the cmd prompt icon <img src="images/cmd.png" width="25" height="25"> in the taskbar or typing "command prompt" in the search field in the taskbar.
2. Type/copy the following command in the command prompt (note that I modified the name of the file at the end of the command to match the file that we have (i.e., "rclone_custom.conf" to "rclone_custom_prod.conf"): </br>
C:\Windows\rclone-v1.68.1-windows-amd64\rclone mount echoview-union: F: -o
UserName=pam_user --vfs-cache-mode full --vfs-cache-max-size 50G --file-perms 0777
--dir-perms 0777 --network-mode --config=C:\Users\pam_user\Desktop\rclone_custom_prod.conf
   1. You will see "The service rclone has been started" in the command prompt and you should see the "F:" drive mounted as "echoview-union" in your Windows Explorer.
   2. Essentially, the F drive is a "virtual mirror" of the GCS prod bucket. This provides a couple of advantages and disadvantages.
      1. **Use the F drive to add files to a EV file.** This is important! If you do not, you are taking advantage of this feature and your read time will be extremely slow.
      2. The .evi file will be written to the data directory on the F drive. But remember this is a virtual mirror and the .evi file does not exist anywhere else.
      3. You can save the EV file anywhere. We recommend not saving the EV file to the F drive, because, again, it exists in a virtual world. We are setting up the directory structure so that the EV files can be written to the GCS prod bucket in a way that does not interfere with the data structure. Stay tuned for more information.
      4. After the .evi file(s) have been written to the F drive, you will want to **copy/move them to the GCS prod data directory** where your files really reside. This is really important to have a permanent copy of the .evi files.
         1. Using the Windows Explorer, copy/move the .evi files from the F drive to the data drive on the Q drive (which should point to the prod bucket - see [Link to GCS prod Bucket](#Link-to-GCS-prod-Bucket) above.
1. Now, you can read your data files directly from the GCS prod bucket without using the F virtual mirror - i.e., you can close the F mount, read the data files directly from the GCS prod bucket, and your read times should be normal for files over a network.

###This is a test again


