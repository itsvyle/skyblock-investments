# skyblock-investments
A google apps script to track skyblock investments over time dynamically through getting prices from api, etc

Template to copy: https://docs.google.com/spreadsheets/d/11dOQrzW2we7vC8Mp2HzTtlGsxgpWaJkz2pTt9K_-evM/ 

## Installing
- Click on `App Scripts` in the extension tab
![image](https://github.com/user-attachments/assets/85860f3e-8d00-4715-9c1a-9dd8f3c7cfad)
- Paste the contents of the `Code.gs` file, and rename the script to "Skyblock Investements", or something along these lines. Save the file with `Ctrl+S`
![image](https://github.com/user-attachments/assets/b0b3508f-908c-4860-aa0b-88a34bc6ea33)
- Going back to the sheet, reload the page, and run the script by going to extensions. Give all asked authorizations - this will only apply to your account and this script, this is not granting me any access
![image](https://github.com/user-attachments/assets/3ba56df3-1f20-4b8d-bd50-a0e8368130a6)
- Simply run the script! You can also set it up so it runs periodically with app scripts triggers from google!


## Usage guidelines
- Fill in your items in the `items` sheet
- The item name must be a valid ID; if it is a rune, it should be `UNIQUE_RUNE.<rune ID>.<rune level>` (eg: `UNIQUE_RUNE.RAINY_DAY.3`)
- Never delete an item - simply change the sold quantity to be the same as buy quantity
- There can be no duplicate IDs; should you decide to buy more of an item at another price, average the prices out

## Widget
You can also get an IOS widget to get the data, by deploying the app script, getting the URL and putting it in the [widget](https://github.com/itsvyle/skyblock-investments/blob/main/widget.powerwidget) file, and then using [Powerwidget](https://apps.apple.com/us/app/power-widgets/id1545771094) on your IOS devices

<img src="https://github.com/user-attachments/assets/4fed2695-5871-4471-b462-76013b511122" height="500">
