import java.util.Scanner;

public class BowlingTeam {

	Scanner scnr = new Scanner(System.in);
	
	int teamID = 0;
	int gamesWon = 0;
	int gamesLost = 0;
	int teamHandicap = 0;
	int teamCurrAverage = 0;
	int teamEntryAverage = 0;
	int scratchTotPinFall = 0;
	int handicapTotPinFall = 0;
	
	int weeksInSeason = 12;
	
	String TeamName = "";
	
	int currentWeekLaneAssignment = 0;
	int nextWeekLaneAssignment = 0;
	int weeklyAssignments[] = new int[weeksInSeason];
	
	int prevWeekGamesWon = 0;
	int prevWeekGameOnePinFall = 0;
	int prevWeekGameTwoPinFall = 0;
	int prevWeekGameThreePinFall = 0;
	int prevWeekTotalPinFall = 0;
	
	
	public BowlingTeam() { // constructor
		teamID = getTeamID();
		
		
	}
	
	
	int getTeamID() {
		int ID = -1;
		
		System.out.println("Please enter the team number:");
		ID = scnr.nextInt();
		return ID;
	}
	
}

/*	Team Class
 * 
 * Team Number(id)
 * Games Won
 * Games Lost
 * Team Handicap
 * Team current average
 * Team Entry average
 * Scratch total pin fall
 * Total pin fall with handicap
 * 
 * Team Name
 * 
 * Current week lane assignment
 * Next week lane assignment
 * array with all weekly assignments set
 * 
 * Previous week games won
 * Previous week game 1 pin fall
 * Previous week game 2 pin fall
 * Previous week game 3 pin fall
 * Previous week total pin fall
 */
