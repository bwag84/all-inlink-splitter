Context
Attached is a demo file of a screaming frog crawl of fedex.com. It contains a sample of all internal links. The master file does indeed contain all links, making it over 900 MB large. Which means most people cannot open it on their laptop.

What I want to create is a reusable Python script that splits this type of file into usable chunks

I want to be able to split files on URL pattern but also on geographical region. We can use the column source segment for that.

Please take into account that because the file is so large it has 13 tabs. That 13 might be variable so it might be less might be more depending on the crawl take that into account. All this material needs to be ingested.

We have built something similar before, but that was bespoke to broken internal links. Those files are a lot less large, but you can use it as a starting point see attached. See splitter.py.

Output
A command line interface tool for Mac that runs a Python
It has an input folder
From that input folder, it will read Excel files
Parse those Excel files
Place the results in the output folder
During processing it will in the command line indicate what step it is on and how many of the how many steps this is so four out of five for example
Create a git repository
Commit and push to that repository independently to my github, let me know if you need to set up permissions (https://github.com/bwag84)
Create a readme.MD that doubles as a manual for this tool
Dockerize the tool and create Dockerfile and run.sh
