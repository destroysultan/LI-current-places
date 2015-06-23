# LI-current-places
Grabs some data from some places!

Setup:
  
    npm install 

Open up the node_modules folder, open the workbook folder, open up workbook.js.
   
    Replace 
    
              if (cell.v == null) continue;
              
    with:
    
              if (cell == null || cell.v == null) continue;

Save and continue.

Create a file with a list of URLs
In app.js, currently line 41ish, replace the fileName with whatever file you wanna use. 
It'll be easier if it's in the same directory.


In the command line just run:
    
    foreman start


Bunch of things will happen, you'll see people's names scrolling, when it stops and says DONNNNNEE then we are DONE!

Your results should be in li-positions.xlsx

Have fun!
