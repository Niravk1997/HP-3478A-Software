# HP 3478A Control and Data Logging Software
 This software allows you to control and log data from the HP 3478A multimeter via an AR488 Arduino GPIB Adapter. This software supports Window 10, 8, and 7.

#### [Download](https://github.com/Niravk1997/DC-Power-Efficiency-Measurement-Software/releases)

#### [User Manual](https://github.com/Niravk1997/DC-Power-Efficiency-Measurement-Software/releases)

#### [AR488 Adapter](https://github.com/Twilight-Logic/AR488)

**Features:**

- Control and Log data, save data into organized folders
- Multithreading Support:
   - All Windows open in a new thread, this ensures maximum performance. For example: Interacting with the Graph Window does not slow down other Windows.
    - All Serial communication happens on an excusive thread. This allows the software to maintain maximum sample capture speed at all times, regardless of what other data processing might be going on.
    - Users can interact with the Graph Window smoothly without any lag.
    - Data Logging functions also run on their designated thread, periodically the software will save measurement data from FIFO data structures into text files.
- Speech Synthesizer feature allows the software to voice measurements periodically and or when it meets the maximum or minimum value threshold.
- Graph Window allows users to visualize their captured data. You can get statistics for all the samples capture or for select few samples. Pan, zoom, and zoom to highlighted area. Save/copy graph as image or save graph's data into text/csv files. 
- Create math waveforms from the samples captured. Create math waveforms from math waveforms. There is no limit to how many math waveforms you can create. 
- Create Histogram from the samples captured and from math waveforms. There is no limit to how many histogram waveforms you can create.
- Measurement table allows users to collect and display the measurement data into a table.
- Capture up to 107 measurement samples in 2 seconds. The fastest sample capture speed compared to any other HP 3478A GUI software.

##### The main software window
![HP 3478A Software](C:\Users\nirav\Documents\GitHub\HP-3478A-Software\Images\HP 3478A Software.gif)


##### Interactive Graphing Module
![HP 3478A Graph Module](C:\Users\nirav\Documents\GitHub\HP-3478A-Software\Images\Graph Module.gif)

##### Create Math Waveform
![HP 3478A Math Wavform](C:\Users\nirav\Documents\GitHub\HP-3478A-Software\Images\Math Waveform.gif)

##### Histogram Waveform
![HP 3478A Histogram](C:\Users\nirav\Documents\GitHub\HP-3478A-Software\Images\Histogram.gif)

##### Measurement Table
![HP 3478A Table](C:\Users\nirav\Documents\GitHub\HP-3478A-Software\Images\Table.gif)

##### How to Connect:
![HP 3478A Connect](C:\Users\nirav\Documents\GitHub\HP-3478A-Software\Images\How to Connect.gif)