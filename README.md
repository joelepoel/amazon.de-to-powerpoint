AMAZON TO POWERPOINT

Amazon to PowerPoint is an automatic PowerPoint generator that converts Amazon product pages into a formatted presentation. Each product gets its own slide with essential information such as product image, title, price, and specifications. This project allows you to generate a PowerPoint presentation without needing to own Microsoft PowerPoint.
Before You Start (One-Time Setup)

    Set up headers.py
        Open headers.py and follow the instructions to set up your user agent.
        
    Install Dependencies
    
        Run the following command to install the necessary Python dependencies:
        pip install -r requirements.txt

How to Personalize

    Open project.py and go to lines 39-43.
    Modify the values on those lines to match your preferences (e.g., template settings, text formatting, etc.).
    Save the changes to project.py.

How to Use

    Ensure that the above one-time setup and personalization steps are complete.
    Open links.txt and paste the URLs of the Amazon products you want to include in the presentation (one URL per line).
    Save links.txt.
    (Optional) If you want to include a logo in your presentation, place an image named logo.png in the project folder.
    Run project.py and provide a filename for the PowerPoint presentation when prompted.
    That's it! Your PowerPoint presentation will be generated and saved with the given filename.

Detailed Description

Amazon to PowerPoint reads product URLs from a list and outputs a PowerPoint presentation with a title slide and a slide for each product.
Core Functions in project.py:

    main function:
    Prompts the user for a filename and coordinates the execution of the other functions. All customization settings are centralized here for easy modification.

    linkreader function:
    Reads the product URLs from links.txt and adds them to a list for processing.

    linkchecker function:
    Validates the URLs to ensure they are valid Amazon.de product links. If any link is invalid, the program exits.

    scraper function:
    Scrapes the specified product URLs for the product’s title, price, specifications, and image. The image is downloaded and saved in the project folder.

    presentationer function:
    Creates the PowerPoint presentation with:
        A title slide
        A slide for each product
        A logo (if provided)
        Saves the presentation with the name you enter.

    imageremover function:
    Cleans up the folder by removing the product images that were downloaded during the scraping process, ensuring there are no leftover files.

Project Structure

.
├── headers.py          # Contains user agent setup
├── links.txt           # List of Amazon product URLs (one per line)
├── logo.png (optional) # Logo image to be included in PowerPoint
├── project.py          # Main script to run the program
├── requirements.txt    # List of dependencies
└── README.md           # This file

Dependencies

The project uses the following Python libraries:

    requests
    beautifulsoup4
    python-pptx
    Pillow
    And others specified in requirements.txt.

Make sure to install the dependencies with the following command:

pip install -r requirements.txt

Notes

    This project is currently designed to work with Amazon.de product links.
    Ensure you have the required permissions to scrape Amazon or any website in accordance with their terms of service.
    The project is designed to make a simple PowerPoint presentation, and you can customize it further by modifying the code.

License

This project is open-source. You may freely use, modify, and distribute it with attribution. For commercial use, make sure to verify it complies with Amazon’s terms of service.
Conclusion

With this tool, you can easily create a PowerPoint presentation from product information on Amazon. Whether you're gathering product information for a presentation or just want to automate this process, this tool can save you time.