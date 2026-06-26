# AA-SI_Echoview
The AA-SI currently is using a timed cloud license on a version of the PAM Windows VM. This license has three seats, meaning three people can be using Echoview simultaneously. We hope to have an annual license soon. The cloud-licensing instructions are found in the PDF document, "Echoview cloud licensing instructions.pdf" that can be accessed via: "P:\Echoview\Echoview cloud licensing instructions.pdf" (on the "dev" GCS environment) or here [Echoview Cloud License](./docs/Echoview_Echoview-cloud-licensing-instructions.pdf). If you run Echoview and the license is not found or working, please see these instructions. </br></br>

Echoview essentially works the same on the AA-SI Windows VM as it does on a local PC, with one exception (that we've found so far) - that is the creation of .evi files. See the section [evi files](#evi-files) for more information. 

# evi Files
When Echoview first reads a data file, or when the version of Echoview has changed such that the format of the .evi file changes, Echoview creates a .evi file. This file is used by Echoview to read a data file more efficiently and faster than an initial read. 


