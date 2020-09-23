<div align="center">

## IPVoiceChat


</div>

### Description

A DirectX method For Speaking over the internet. I am having some problems with perfecting it, but if you like what you see let me know.

This code has evolved allot over the past few months. First I tried to design it to use straight api with a wave format, but the data transfer of such a format was not practical because the wave header format would be change with every transfer. DirectX was the best way to go because I just read the buffer from the client PC and wrote it to the receiving PC. There is allot of functionality in the app that isn't being used for the above reasons
 
### More Info
 
Some of the problems I am having is:

1. The data isn't in real time. you lose about a half second of sound from the originating PC and there is about a one second skip on the playing PC as the Sound Buffer is written with the new data.

2. The way that the data is written into the sound buffer is not right. At this point I destroy and recreate the buffer with every new transfer of data, I have had little success with creating a looping buffer that writes new data into the buffer dynamically. All the DirectX documentation says you can do this but this app has special characteristics that make it very difficult, such as the network speed makes it difficult to determine when a new data transfer will arrive.

The way the app works is when on a network the app must be running on two PCs (the future design may handle more connections) from either PC type in the the DNS or IP address of the computer you want to connect to and press enter. If you are connecting the other PC will flash a message box asking if you wish to except the connection, say yes and you are connected. Click the Talk button to start the recording from the Microphone.

Hopes This helps I have be working on this for about 4 months, and it's starting the make me really mad. So please stay in touch with your changes.

Thanks Allot


<span>             |<span>
---                |---
**Submitted On**   |2001-11-01 21:23:16
**By**             |[James Thomas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-thomas.md)
**Level**          |Advanced
**User Rating**    |3.7 (37 globes from 10 users)
**Compatibility**  |VB 6\.0
**Category**       |[DirectX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/directx__1-44.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[IPVoiceCha328241122001\.zip](https://github.com/Planet-Source-Code/james-thomas-ipvoicechat__1-28590/archive/master.zip)








