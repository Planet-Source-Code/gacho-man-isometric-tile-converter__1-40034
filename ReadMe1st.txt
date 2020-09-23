--- Daniel Marin's ISOMETRIC CONVERTER ---

Hi there, this utility is for my coming Isometric-tile-based new game

This is the first time I share code in a site, so, I hope my code is useful to all of you

all the comments can be sent to: marindaniel@yahoo.com

--- Instructions for the program ---

If you have questions, I wrote here the FAQ:

    -What is this SH** I downloaded?

      * Is an utility that transforms rectangular images to isometric ones, the isometric tiles 
        have a specific structure so they can be arranged as tiles in a tiled game, with these
        simple formulae:

               The Coordinate X = (TileX + TileY) * TileWidth / 2
               The Coordinate Y = [(MapHeight / 2) + (TileY - TileY)] * TileHeight / 2

    -How does this MO**** F***ING thing works?

      * The main function is in the "Transform to Isometric" button, for that, you need to select
        in what you want to transform it, in a floor tile, or in a right or left wall, these
        options are un the inferior-right section of the main window, also with the backcolor
        of the picturebox for the isometric result, and finally, select an image, for this
        there are two ways: Selection in the superior-left section of the main window, an image
        file, automatically the program will open that file and show it to you, the other way
        is to paste a bitmap image from the clipboard (with the typical CTRL+V). When the 
	transformation ends, you can copy it to the clipboard (with the hypertypical CTRL+C)

    -I don't understand your F***ING code dude!

      * I'm not very used to use a lot of comments, I tried to use the necesary ones, for the
        total understanding of the code you need to be a basic-intermediate visual basic
        programmer/understander, and a little of maths and advanced artificial intelligence
        algorythmics for autopatterns creation (LOL)

    -Your program sucks!!!!

      * Create your own and send it to me, I always can learn

    -You son of a B****!!! my question was not here!!!

      * OK... ask me at marindaniel@yahoo.com