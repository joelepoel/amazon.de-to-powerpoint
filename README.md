# AMAZON TO POWERPOINT

#### Description:

    In short:

    This code allows you to generate a powerpoint with a title slide and one slide per product found on Amazon.
    Every slide will be complete with an image, price, title and the specifications of the specified product.
    You don't need to own Powerpoint to use it.



    Before you use(one-time steps):
    -Open headers.py and follow short instructions (user agent)
    -pip install the listed dependencies in requirements.txt (run: pip install -r requirements)



    How to personalize:
    1. in project.py, on lines 39-43, change the values to your liking
    2. save project.py



    How to use:
    1. Do the steps above to make sure the code works and is personalized to your liking
    2. Open links.txt
    3. Paste the urls of the products you want to use in links.txt (one per line)
    4. Save links.txt
    5. If you want to use a logo, add a logo to the folder with the name "logo.png"
    6. Run project.py and pick a filename
    7. Ta-da!



    A more detailed description of the project:

    AMAZON TO POWERPOINT is an automatic Powerpoint generator with images, productspecifications and the titles from products sold on Amazon.de
    reads from a list of product urls and outputs them into a Powerpoint presentation.

    Project.py exists out of a few functions:
    - The "main" function prompts the user for a filename, and calls all the other functions. I have also put all the personalization variables here, to make it easier for the user.
    - The "linkreader" function reads all urls in the linkts.txt file and appends them to a list
    - The "linkchecker" function checks the validity of the given urls. If a link is not a valid amazon.de link, it gives a sytem exit.
    - The "scraper" function scrapes the given urls for a product title, price and productspecifications. It also saves the product image to the folder.
    - The "presentationer" function creates a Powerpoint presentation with a title slide and creates a separate slide for each scraped product. It also adds a logo, if given, and it saves the Powerpoint presentation as givenname.pptx, wherein givenname acts as the name you entered.
    - The "imageremover" function removes all the images, that were scraped from the urls, from the folder. This way there are no leftovers and you don't have to worry about that when you use the code another time.
"# amazon.de-to-powerpoint" 
"# amazon.de-to-powerpoint" 
