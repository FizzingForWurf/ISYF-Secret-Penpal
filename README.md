# ISYF Secret Penpal

## Python program to allocate pairs for the Secret Penpal activity.

Secret penpal is an activity where each participant will be writing to another participant whose identity will be unknown to them. Each pair will write to each other about their experience at the end of each day of the event. The identity of one's penpal will only be revealed at the end of the event. 

Participants will be paired up according to these criteria: 
- Both participants are from different **groups** 
- Both participants are from different **schools**
- Foreign students will **always** be paired with locals 

This is to promote culture appreciation among the participants and reduce the chances for participants to find their penpal during the events before the reveal.

Note: This program is written for the Secret Penpal activity conducted during the International Science Youth Forum (ISYF). 

## Download
For Windows, please download the executable file in the `windows` folder.

For Mac, please download the executable file in the `mac` folder.

## How to use
For the program to interpret the participants information correctly, please follow the format as shown below strictly.

1. The **1<sup>st</sup> sheet** in the input file should contain all **LOCAL** schools. It should start at the top right cell. Please do NOT include other information other than local schools in the first column.

2. The **2<sup>nd</sup> sheet** should contain the delegates information in their respective groups.

    1. The **1<sup>st</sup> ROW** contains the **headers** (Name, School, Nationality, Gender). This row of information will be ignored by the program

    2. **1<sup>st</sup> COLUMN**: Name of participant
    3. **2<sup>nd</sup> COLUMN**: School of participant
    4. **3<sup>rd</sup> COLUMN**: Nationality of participant
    5. **4<sup>th</sup> COLUMN**: Gender of participant

    > *Important*: Leave an **EMPTY ROW** between participants of **different** groups. This is for the program to identify the groupings

3. Any other sheets should be placed **BEHIND** the first two abovementioned sheets. These additional sheets will be ignored by the program.

4. Ensure that the input excel sheet is in the **SAME FOLDER** as the program so that it can find it! Next, input the name of the excel sheet (without `.xlsx`)

5. Lastly, input the number of participants and number of groups in the event for verification purposes.

## Troubleshooting
1. Pair cannot be found 

    - If the output repeatedly shows one participant's name, simply terminate the program and run it again.

2. Incorrect number of groups or participants

    - Likely a formatting issue with the input excel sheet. Please ensure it follows the abovementioned formatting requirements.

3. Other issues

    - Contact me at my Telegram handle @zheng_hongg or email me at zhtong@gmail.com for any issues. I will do my best to help! :relaxed:


## Implementation
Since there are more foreign students than local students, this program will prioritise matching foreign students first. When matching a student with another person, another participant will be **randomly** selected. 

The selection criteria as mentioned above will be checked between the two participants. If it is a valid pair, they will be recorded and removed from the pool of participants. Else, anther random student will be selected. 

This matching process repeats until there is either one or no participants left. If the total number of participants is odd, there will be an additional person left. This person will be need to be **manually added** to a pair!

## Find a bug?
If you found an issue or would like to submit an improvement to this project, please submit an issue using the issues tab above. If you would like to submit a PR with a fix, reference the issue you created!