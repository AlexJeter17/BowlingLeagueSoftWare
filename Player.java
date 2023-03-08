import java.util.Scanner;

public class Player {
	int teamID, 
	average, 
	entryAverage, 
	totGamesPlayed,
	totalPinFall, 
	gender; // Male = 1 || Female = 0

	double handicap = 0.0;
	
	Scanner scnr = new Scanner(System.in);
	
	public Player() {
		average = 0;
		entryAverage = 0;
		handicap = ((230 - average) * 0.8);
		totGamesPlayed = 0;
		totalPinFall = 0;
		gender = -1;
	}
	
	
	public void getGender() { // Gets the Gender of a player and will loop until getting a valid number
		while(gender != 0 && gender != 1) {
		System.out.println("Please type in 0 for Female Player or 1 for Male Player and hit 'Enter'");	
		gender = scnr.nextInt();
		}
	}
}

/*  Player Class
 * 
 * Team ID
 * Average
 * Entry Average
 * Handicap (80% of 230)
 * Total Games played
 * Total pin fall
 * Gender
 * 
 */
