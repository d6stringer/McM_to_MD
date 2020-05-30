# bm_repo

Created by: Daniel Woodson<br />
In May, 2020<br /><br />

Purpose: Making purchase requests for McMaster was getting tedious with the copipasta and formatting for markup
    so that it was easy to use seemed like a logical next step.<br /><br />

Workflow:<br />
    1. Make your McMaster cart like you normally would.<br />
    2. Email the McM cart to an Outlook account you have access to.<br />
    3. Save/Drag the message (*.msg format) to a location on your machine where you can find it.<br />
        I would use your desktop.<br />
    4. Run the executable and find your file.<br />
    5. Once you select your file and open it, the program will automagically copy the correct info from the
        file to your clipboard.<br />
    6. Create a new purchase card in trello.<br />
    7. Paste the copied data to your new trello card and save the card.<br />
    8. Wait with anticipation for your parts to arrive.<br />
    9. Rinse and repeat.<br /><br />
Executable can be found in the dist folder.<br /><br />
Executable created with pyinstaller:<br />
    from command line:<br />
    pyinstaller --onefile main.py<br />

