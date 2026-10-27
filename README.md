# Games For Sheets!

A few classic games recreated on Google Sheets! I might add more games, but the ones currently on there are Snake, Blackjack, Minesweeper, Sudoku, and Tetris. It was really fun making all of these, and I hope you have fun playing them too! (or at least play the blackjack in school or smthn)

<img width="883" height="624" alt="image" src="https://github.com/user-attachments/assets/c24555b4-3f4e-488d-9244-dd969fd7b46e" />

## Features

- Google Sheets recreations of Minesweeper, Snake, Blackjack, and Sudoku.
- Made with Google Apps Script (basically JavaScript for Google applications)


## Requirements

- A good, stable internet connection
- A powerful enough device to load Google Sheets
- A browser

## Use

If you want to try it out for yourself, go to [this sheet](https://docs.google.com/spreadsheets/d/14v2c1DsDYz6Efg_-XwYdSOKMCCL0Xanr9LB-ZKnGtPw/edit?gid=0#gid=0), and make a copy of it by going to __File -> Make a Copy -> Make a Copy__. Then you can select the game you want to play from the bottom and press Start. How to play the different games is realy game specific, so there's a little snippet of text to help you get on the right track to start playing it.

Also, if you are going to use the [xlsx](https://github.com/CHeesee69/Games-For-Sheets/blob/main/Games_for_sheets.xlsx), you are going to have to replace all of the cells that say "FALSE" with checkboxes that act as some of the game's controls. You will also have to connect the start buttons to their functions inside the scripts. If you have no idea what that means, just use the method of copying the sheet.

### Blackjack

<img width="1438" height="792" alt="image" src="https://github.com/user-attachments/assets/e4772fb2-9fb4-4884-8db6-66d01e06cd40" />

Blackjack is a popular casino game where the goal is to get your cards closest to 21. On your turn, you can either hit or stand. If you hit, the dealer will give you another card. If the sum of all of your cards (number cards are worth their number, face cards are worth 10, and aces are worth 1 or 11) is over 21, you bust. If its close, you might want to stand. The version that is on the sheet is one-on-one with the dealer to see who can get a higher amount.

 ### Snake

<img width="1598" height="722" alt="image" src="https://github.com/user-attachments/assets/c4eef27f-2c45-41f7-aa2c-beaf1477ea88" />

Snake is a game that is you controling a snake in a grid that goes around eating apples. With each apple that is eaten, the snake gets longer. If you crash into the wall or your own body, the game ends. On the sheet, I made two versions of it: a 15 x 15, and a 10 x 10. This was what this project was originaly so it has the most features and customisation.

### Minesweeper

<img width="823" height="450" alt="image" src="https://github.com/user-attachments/assets/8213c57d-5016-466a-b62a-5f42c454bcd9" />

Minesweeper is a game about looking for a specific amout of mines in a minefeild. When you dig on a square, it will either be empty, have a number, or have a mine underneath it. The number telly you exactly how many mines are in the adjacent 8 squares. So if its a 1, there is only one mine in that 3 x 3 area around that 1. I made a 10 x 10 version of it with the option to increace the mines to up to 25, or decreace them to 10.

### Sudoku

<img width="1140" height="633" alt="Screenshot 2026-10-26 132851" src="https://github.com/user-attachments/assets/40759005-f43c-4f9d-aff7-c3a6939fc64d" />

Sudoku is a game that consists of a 9 x 9 grid with a number in some of the squares. each of the 9 3 x 3 boxes must have each number 1 - 9. But, only one of each number are allowed on any given collum or row. The game finishes when you fill up all of the numbers in the 9 x 9 grid. The one in the sheets has a easy, medium, and hard difficulty. The easy difficulty starts you out with more numbers, allowing you to more easaly spot missing ones. The hard one has less, making it harder to find exactly where each number goes.

### Tetris

<img width="1068" height="728" alt="image" src="https://github.com/user-attachments/assets/91b6005e-6e68-4aa5-9522-66ac2e8f2742" />

Tetris is a well known classic game that was popularised when it was bundled in with the Gameboy. Diffrent shaped blocks fall from the top and you have to figure out a place to put them. if you fill a row, that row dissapears. The more rows you clear at a time, the more points you get. if you clear 4 rows with a four piece, you get a tetris. If the new piece at the top intersects with an existing piece, you lose. The game is 10 x 20 spaces and can go on as long as you can keep up with the randomness of the game.
