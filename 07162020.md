# 7/16/2020 - Week 3
## Clean up time for the games list

## Clean up items:
### ESRBWhy Cleanup
* VBA --> Perform melting on ESRBWhy column's values. separate out the values within, to have their own columns for each possible type. Steps taken, below:
  1. Copy all values to new tab, use "Text to Columns" feature to push values out, delimited by Pipes (|).
  2. Write some VBA to copy values and place them into a giant concatenated list.
  3. Remove Duplicates from list, remove empty row(s), then clean up values in lists, using a custom function in VBA combined with TRIM function from Excel.
  4. Remove Duplicates again, and review.
  * --- In reviewing, the delimitation from the "Text to Columns" usage did not separate out commas into their own columns as well. Going to redo that step, with commas as delimiters, then rerun the above steps.
  5. REDO with commas in delimiters settings!! Also with semicolons... You know what.... ALL THE DELIMITERS EXCEPT SPACES!!
  6. In further reviewing, post changing delimiters, there are some fields where the delimitation is actually being used as such, and others where it isn't. I believe the best course of action, at this point, is to review the manually created ~250 rows for themes, and create searches for common words instead. Flagging the rows this way, seems more accurate in the long run.
      1. Replace link words/characters, like "and" or "&," with delimiters, and separate them out into a list of keywords that can be used as a flag counter of sorts
      2. Created final list of keywords, and got counts on whether they were found in each possible entry.
      3. Decided to Keep anything that had over 100 mentions from the ESRBWhy column, removing unnecessary values that were repetitive, vague, or adverbs/adjectives/etc. 
      4. 17 values remain. Will turn these into columns, that have a Y/N if the word exists in the ESRBWhy column.

### Game_ID Cleanup
* Remove excess entries from Game_ID list. Get rid of lists, keeping only the first game id. The VBA code followed this procedure when continuing the scraping.
  1. Used if statement to check for a comma being present or not. This dictated excess entries. Chose first entry using left/mid functions. Then replaced all Game_ID's with first entry only.
  2. Removed duplicates across board on the file. Checked for more dups, post remove duplicates operation to verify. 202 values still found. This means that the remaining columns from the file, may have information that was different, based on some unknown reason.
  3. On researching the remaining 202 values, the reason was usually 1 of a few things.
      1. The website had the text "sub" instead of "app" in the URL (Site column) at a specific location.
      2. The total reviews changed LIVE, at some point during the webscraping, causing the "Total_Reviews" column to differ by 1.
      3. The bundles sometimes had a different Game_ID used in the URL, presumably just a different game ID, or a specific ID for the bundle.
  4. In order to preserve the row placement, I decided to remove entries that were NOT the first entry of the game. The factors above dignify that the game may have been better sold in a bundle, rather than on their own, but the complexities of decision making for this wound up seeming to be of less value than simply capturing the game in its "highest" position in the list. The order of the games was kept throughout this process, and no sorting was done, up to this point. To do the removals in this method, I just used a COUNTIF function in Excel, where there was an absolute reference to the first cell at A2. This means, that anything with a count of Game_ID higher than 1, had already appeared in the list and could be removed. We're down to 15,824 entries now, after removing duplicate Game_ID's.
  
### Release Date Cleanup
* Discern release date, based off of whether or not Release1 (R1) or Release2 (R2) has a valid date. Determine a method to settle disagreements between the values in each, if they occur.
  1. Determine how to tell, systematically, if a date in a field is a date or not, in field format/type, to Excel. This was done using an ISERROR function in Excel, enclosed around a multiplication, of the cell in question, by 1. If the value was a date, the result would be FALSE, if not a date to Excel, would return TRUE.
  2. Create a means to change NON-dates into dates. Verify accuracy.
  3. Create in IF check. Checks if date in field is a date or not. If it is, pull date, else, convert to date.
  4. For some reason, dates that are text based, are NOT into double digits on the day number. None of the dates are at or beyond the 10th of any month...
  5. After applying logic from steps above, the dates that were NOT able to be fixed, wound up totally to only 31 in count. All of which fell into 2 categories:
      1. Had No date given in either Release1 or Release2 columns. i.e. they said "Spring 2021"
      2. Had a Release2 date, but NOT a Release1 date that actually had a complete date.
      3. Had a Release1 date of 2020, but a Release2 date of 7/12/1905 for some reason.
  6. Fixed rows that had a bad Release1, to use Release2 instead, where able. This left me with only 6 rows with a definite bad date value. I will handle these after checking for R1 vs R2 date disagreements next.
  7. In checking for date disagreements, only 1 row came up as a true "issue" for this reason. It looks to be a typo on someone's part, as R1 = 8/28/2002 and R2 = 8/28/2020. In looking at the row in question, the game was "Wasteland 3" which I know to be a newer game, so the 2020 date is correct. I manually forced the entry on this field's resulting value. The remaining rows that had a disagreement were due to R2 being blank. Otherwise R1 was always = R2, when the dates were actually pulled as date values. Anything where the date value in R1 was listed as a text value, I compared the conversion for, with R2. The results were exactly the same as the comparison of R1 to R2 again. So, Even if R1 was a text value, it ALWAYS matched R2.
  8. Last step, handle the oddities in dates given. Oddities Encountered, and solution arrived upon:
      * R1 = 2020 & R2 = 7/12/1905 - Looked these up manually. They are all for games that are NOT released yet. DELETE ROW
      * Season YYYY for R1 & R2 - Games that are also Not Released yet. DELETE ROW
      * No R1 or R2 value, 14 found. Conversion registered as 0 value for julian date of 1/1/1900. Manually looked these up on Google to get release dates. Sometimes this was NOT easy to do, because the date the game came out, was NOT the same date that the game may have become available on Steam, rather than a console. OR, the version of the game may have been replaced with the deluxe version ONLY, as the game got older. As an example, Borderlands 1 came out sometime before 2013 (I wanna say 2010 or earlier maybe). However, the "Game of the Year" edition came out at a later date. I opted to delete these rows as well, since we simply just can't be sure of when they came out.
All in all, only removed 20 games due to bad release dates. Not bad!

### Final Thoughts for 7/16/2020:
Final version of the data only has 15805 rows now.
Some additional things I'd like to do, or possibly am thinking of attempting to do:
  1. Update Franchise column/concept. Instead of showing the Franchise, see if we can show a Y/N franchise value, then the count of games IN that Franchise. Probably needs cleanup on the column.
  2. Cleanup Developer/Publisher columns
  3. Cleanup and Split out the Genre tags into their own columns
  4. Cleanup the cost/price column
  5. decide what to do with the Metacritic and Steam Positive ratings. Maybe aggregate them somehow?
  6. Possibly come up with date range buckets for release dates vs. the date the web-scraping was done. Review prices from that perspective.

