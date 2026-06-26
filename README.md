# AA-SI_Echoview
The AA-SI currently is using a timed cloud license on a version of the PAM Windows VM. This license has three seats, meaning three people can be using Echoview simultaneously. We hope to have an annual license soon.  For more information on the Windows VM, see the [AA-SI GitHub Windows VM repo](https://github.com/nmfs-ost/AA-SI_WindowsVM) </br>

The cloud-licensing instructions are found in the PDF document, "Echoview cloud licensing instructions.pdf" that can be accessed via: "P:\Echoview\Echoview cloud licensing instructions.pdf" (on the "dev" GCS environment) or here [Echoview Cloud License](./docs/Echoview_Echoview_cloud_licensing_instructions.pdf). If you run Echoview and the license is not found or working, please see these instructions. </br>

Echoview essentially works the same on the AA-SI Windows VM as it does on a local PC, with one exception (that we've found so far) - that is the creation of .evi files. See the section [evi files](#evi-files) for more information. 

# evi Files
When Echoview first reads a data file, or when the version of Echoview has changed such that the format of the .evi file changes, Echoview creates a .evi file. This file is used by Echoview to read a data file more efficiently and faster than an initial read. These .evi files are by default written to the same directory as the data files. On a local PC this is not a problem. On local networks, this process can slow down initial reads of data when/if there are network interuptions. However, reading data files and writing the .evi files from the Windows VM causes severe limitations. For example, in initial testing we found that it took 25 minutes to read one 200-MB EK60 .raw file! After the .evi file is written, data reading is as fast as on local PCs or networks for subsequent reads. We are working with Echoview to modify how and where they write the .evi files, but in the interim, Daniel Woodrich created an alternative method that works well, but requires a few more steps. </br>

The instructions are provided in this [idx fix document](docs/Echoview_idx_files_issue_ immediate_solution.pdf).






