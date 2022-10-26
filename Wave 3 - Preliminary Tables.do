********************************************************************************
* PRELIMINARY DESCRIPTIVE TABLES
********************************************************************************
cap log close 

// BEFORE PROCEEDING, RUN MAIN SET-UP DO FILE FIRST!!

log using "$logs/Wave 3 - Oct 2022/Cleaning/Wave 3 Preliminary Tables_`c(current_date)'.smcl", replace

local tranche = $tranche

* open dataset
if $tranche <4 {
use "$data/Wave 3 - Oct 2022/Cleaned/Wave 3_Tranche $tranche.dta", clear
}

if $tranche == 4 {
use "$data/Wave 3 - Oct 2022/FINAL/Wave 3.dta", clear
}


********************************************************************************
* SECTION 2 - INDIVIDUAL-LEVEL DATA
*** SECTION 2.1 - MH SCALES
*** SECTION 2.2 - NON-MH SECTIONS
* SECTION 3 - HH-LEVEL DATA
* SECTION 4 - COMBINE TABLES INTO ONE EXCEL SHEET
********************************************************************************


********************************************************************************
* SECTION 2 - INDIVIDUAL-LEVEL DATA
********************************************************************************

* OVERALL SAMPLE
putexcel set ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Preliminary Tables_Tranche `tranche'_`c(current_date)'.xlsx", ///
sheet("Sample_Tranche`tranche'") replace

putexcel A1 = "Current Sample by Age Group"
putexcel (A1:C1), merge hcenter vcenter

putexcel B2 = "N (Live/Immigrated)"
putexcel (B2:C2), merge hcenter vcenter
putexcel B3 = "Female", hcenter
putexcel C3 = "Male", hcenter
putexcel A4 = "Ages 0-5"
putexcel A5 = "Ages 6-7"
putexcel A6 = "Ages 8-10"
putexcel A7 = "Ages 11-13"
putexcel A8 = "Ages 14-16"
putexcel A9 = "Ages 17-19"
putexcel A10 = "Ages 20-39"
putexcel A11 = "Ages 40-59"
putexcel A12 = "Ages 60+"
putexcel A13 = "Total"

tab agegroup gender, matcell(cellcounts)
putexcel B4 = matrix(cellcounts), hcenter

count if gender=="F" & age_yrs!=.
putexcel B13 = `r(N)', hcenter
count if gender=="M" & age_yrs!=.
putexcel C13 = `r(N)', hcenter

putexcel B15 = "N (Not in HH)", hcenter
putexcel A16 = "Dead"
putexcel A17 = "Emigrated" 
count if member_status=="Dead"
putexcel B16 = `r(N)', hcenter
count if member_status=="Emigrated"
putexcel B17 = `r(N)', hcenter
putexcel A18 = "Total"
count if !inlist(member_status, "Member of the household", "Immigrated")
putexcel B18 = 	`r(N)', hcenter


*****************************************
*** SECTION 2.1 - MH SCALES
*****************************************
global MH_scores RCADSselfreportsum SCASselfreportsum PSCparentreportsum ///
SCASpreschoolparentreportsum SCASschoolageparentreportsum PHQ9sum GAD7sum

global MH_adol RCADSselfreportsum SCASselfreportsum PSCparentreportsum ///
SCASpreschoolparentreportsum SCASschoolageparentreportsum

global MH_elev RCADSselfreportelev SCASselfreportelev PSCparentreportelev SCASpreschoolparentreportelev SCASschoolageparentreportelev PHQ9elev GAD7elev

* Raw scale scores
estpost tabstat $MH_scores, by(gender) stats(N mean sd min max)  c(stats) 

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/MH_RawScaleScore_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace

* Raw scale scores by age group
	putexcel set ///
	"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Preliminary Tables_Tranche `tranche'_`c(current_date)'.xlsx", ///
	sheet("MH_RawScore_AgeGroup") modify
	
	putexcel B1 = "Male"
	putexcel (B1:D1), merge hcenter vcenter
	putexcel B2 = "N"
	putexcel C2 = "Mean"
	putexcel D2 = "SD"
	
	putexcel E1 = "Female"
	putexcel (E1:G1), merge hcenter vcenter
	putexcel E2 = "N"
	putexcel F2 = "Mean"
	putexcel G2 = "SD"
	
	* Scales for Ages 0-5
	putexcel A3 = "Ages 0-5",  border(bottom)
	local i = 4
	foreach var of varlist PSCparentreportsum SCASpreschoolparentreportsum  {
		sum `var' if ///
		agegroup==1 & gender=="M"
		
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)',  nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		
		sum `var' if ///
		agegroup==1 & gender=="F"
		
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 6-7
	putexcel A`i' = "Ages 6-7",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportsum ///
	SCASpreschoolparentreportsum SCASschoolageparentreportsum  {
		
		* Boys
		sum `var' if agegroup==2 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==2 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 8-10
	putexcel A`i' = "Ages 8-10",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportsum SCASschoolageparentreportsum ///
	RCADSselfreportsum SCASselfreportsum  {
		
		* Boys
		sum `var' if agegroup==3 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==3 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 11-13
	putexcel A`i' = "Ages 11-13",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportsum SCASschoolageparentreportsum ///
	RCADSselfreportsum SCASselfreportsum  {
		
		* Boys
		sum `var' if agegroup==4 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==4 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 14-16
	putexcel A`i' = "Ages 14-16",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportsum SCASschoolageparentreportsum ///
	RCADSselfreportsum SCASselfreportsum  {
		
		* Boys
		sum `var' if agegroup==5 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==5 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 17-19
	putexcel A`i' = "Ages 17-19",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportsum SCASschoolageparentreportsum ///
	RCADSselfreportsum SCASselfreportsum  {
		
		* Boys
		sum `var' if agegroup==6 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==6 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 20-39
	putexcel A`i' = "Ages 20-39",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PHQ9sum GAD7sum {
		
		* Men
		sum `var' if agegroup==7 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Women
		sum `var' if agegroup==7 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 40-59
	putexcel A`i' = "Ages 40-59",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PHQ9sum GAD7sum {
		
		* Men
		sum `var' if agegroup==8 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Women
		sum `var' if agegroup==8 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 60+
	putexcel A`i' = "Ages 60+",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PHQ9sum GAD7sum {
		
		* Men
		sum `var' if agegroup==9 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Women
		sum `var' if agegroup==9 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
* Elevated scores
estpost tabstat $MH_elev, by(gender) stats(N mean sd min max)  c(stats) 

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/MH_ElevScale_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace

	putexcel set ///
	"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Preliminary Tables_Tranche `tranche'_`c(current_date)'.xlsx", ///
	sheet("MH_ElevScore_AgeGroup") modify
	
	putexcel B1 = "Male"
	putexcel (B1:D1), merge hcenter vcenter
	putexcel B2 = "N"
	putexcel C2 = "Mean"
	putexcel D2 = "SD"
	
	putexcel E1 = "Female"
	putexcel (E1:G1), merge hcenter vcenter
	putexcel E2 = "N"
	putexcel F2 = "Mean"
	putexcel G2 = "SD"
	
	* Scales for Ages 0-5
	putexcel A3 = "Ages 0-5",  border(bottom)
	local i = 4
	foreach var of varlist PSCparentreportelev SCASpreschoolparentreportelev  {
		sum `var' if ///
		agegroup==1 & gender=="M"
		
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)',  nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		
		sum `var' if ///
		agegroup==1 & gender=="F"
		
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 6-7
	putexcel A`i' = "Ages 6-7",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportelev ///
	SCASpreschoolparentreportelev SCASschoolageparentreportelev  {
		
		* Boys
		sum `var' if agegroup==2 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==2 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 8-10
	putexcel A`i' = "Ages 8-10",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportelev SCASschoolageparentreportelev ///
	RCADSselfreportelev SCASselfreportelev  {
		
		* Boys
		sum `var' if agegroup==3 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==3 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 11-13
	putexcel A`i' = "Ages 11-13",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportelev SCASschoolageparentreportelev ///
	RCADSselfreportelev SCASselfreportelev  {
		
		* Boys
		sum `var' if agegroup==4 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==4 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 14-16
	putexcel A`i' = "Ages 14-16",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportelev SCASschoolageparentreportelev ///
	RCADSselfreportelev SCASselfreportelev  {
		
		* Boys
		sum `var' if agegroup==5 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==5 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 17-19
	putexcel A`i' = "Ages 17-19",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PSCparentreportelev SCASschoolageparentreportelev ///
	RCADSselfreportelev SCASselfreportelev  {
		
		* Boys
		sum `var' if agegroup==6 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==6 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 20-39
	putexcel A`i' = "Ages 20-39",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PHQ9elev GAD7elev {
		
		* Men
		sum `var' if agegroup==7 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Women
		sum `var' if agegroup==7 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 40-59
	putexcel A`i' = "Ages 40-59",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PHQ9elev GAD7elev {
		
		* Men
		sum `var' if agegroup==8 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Women
		sum `var' if agegroup==8 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 60+
	putexcel A`i' = "Ages 60+",  border(bottom)
	local i = `i' + 1
	foreach var of varlist PHQ9elev GAD7elev {
		
		* Men
		sum `var' if agegroup==9 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Women
		sum `var' if agegroup==9 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	




*****************************************
*** SECTION 2.2 - NON-MH SECTIONS
*****************************************

* Time use
global timeuse TimePlayOutdoors TimePlayIndoors TimeAlone TimeInClass TimeStudying TimeHousework

estpost tabstat $timeuse, by(gender) stats(N mean sd min max)  c(stats) 

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/TimeUse_all_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace


	putexcel set ///
	"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Preliminary Tables_Tranche `tranche'_`c(current_date)'.xlsx", ///
	sheet("TimeUse_AgeGroup") modify
	
	putexcel B1 = "Male"
	putexcel (B1:D1), merge hcenter vcenter
	putexcel B2 = "N"
	putexcel C2 = "Mean"
	putexcel D2 = "SD"
	
	putexcel E1 = "Female"
	putexcel (E1:G1), merge hcenter vcenter
	putexcel E2 = "N"
	putexcel F2 = "Mean"
	putexcel G2 = "SD"
	
	* Scales for Ages 0-5
	putexcel A3 = "Age 5",  border(bottom)
	local i = 4
	foreach var of global timeuse  {
		sum `var' if ///
		agegroup==1 & gender=="M"
		
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)',  nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		
		sum `var' if ///
		agegroup==1 & gender=="F"
		
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 6-7
	putexcel A`i' = "Ages 6-7",  border(bottom)
	local i = `i' + 1
	foreach var of global timeuse  {
		
		* Boys
		sum `var' if agegroup==2 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==2 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 8-10
	putexcel A`i' = "Ages 8-10",  border(bottom)
	local i = `i' + 1
	foreach var of global timeuse  {
		
		* Boys
		sum `var' if agegroup==3 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==3 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 11-13
	putexcel A`i' = "Ages 11-13",  border(bottom)
	local i = `i' + 1
	foreach var of global timeuse {
		
		* Boys
		sum `var' if agegroup==4 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==4 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 14-16
	putexcel A`i' = "Ages 14-16",  border(bottom)
	local i = `i' + 1
	foreach var of global timeuse  {
		
		* Boys
		sum `var' if agegroup==5 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==5 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}
	
	* Scales for Ages 17-19
	putexcel A`i' = "Ages 17-19",  border(bottom)
	local i = `i' + 1
	foreach var of global timeuse {
		
		* Boys
		sum `var' if agegroup==6 & gender=="M"
		putexcel A`i' = "`var'"
		putexcel B`i' = `r(N)'
		putexcel C`i' = `r(mean)', nformat(number_d2)
		putexcel D`i' = `r(sd)', nformat(number_d2)
		* Girls
		sum `var' if agegroup==6 & gender=="F"
		putexcel E`i' = `r(N)'
		putexcel F`i' = `r(mean)', nformat(number_d2)
		putexcel G`i' = `r(sd)', nformat(number_d2)
		
		local i = `i'+1
	}


* IND COVID
global ind_covid Ind_CovidHad_yn Ind_CovidSevere_yn Ind_CovidDied_yn

estpost summ $ind_covid 

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/IndCov_all_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace

* HEALTH STATUS
global health_status HealthNeededCare HealthReceivedMedCare DiabetesHistory ///
DiabetesStartRecently DiabetesGetCare

estpost summ $health_status 

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/Health_Status_all_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace

* ADLs
global ADLs DiffRemembering DiffWalking100 DiffClimbing DiffLifting DiffDressing

estpost summ $ADLs 

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/ADLs_all_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace

* IMMUNIZATION & ANC
global vax_anc any_vax_yn ImmuBCG_yn ImmuPolio ImmuDPT ImmuMeasles_yn

	* everyone
	estpost summ $vax_anc if inrange(age_yrs,0,5)

	esttab using ///
	"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/Immu_ANC_all_Tranche `tranche'_`c(current_date)'.csv", ///
	cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
	nomtitle nonote noobs label collabels("N" "Mean" "SD") replace

	* by gender
	estpost tabstat $vax_anc, by(gender) stats(N mean sd min max)  c(stats) 

	esttab using ///
	"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/Immu_ANC_bygender_Tranche `tranche'_`c(current_date)'.csv", ///
	cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
	nomtitle nonote noobs label collabels("N" "Mean" "SD") replace


********************************************************************************
* SECTION 3 - HH-LEVEL DATA
********************************************************************************
eststo clear
preserve
* collapse to hh-level
collapse (first) state region_type HHCov_neighb_had- HHCov_anyone_died agegroup region, by(hh_id)


global hhcov_all HHCov_neighb_had_yn HHCov_neighb_severe_yn HHCov_neighb_died_yn HHCov_fr_rel_had_yn HHCov_fr_rel_severe_yn HHCov_fr_rel_died_yn HHCov_other_had_yn HHCov_other_severe_yn HHCov_other_died_yn

estpost summarize $hhcov_all

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/HHCov_Simple_Summ_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace
	  
	  
	  
* Comparing variables across states
est clear
estpost tabstat  HHCov_anyone_had HHCov_anyone_severe HHCov_anyone_died, ///
by(state) stats(N mean sd) c(stats) listwise

esttab using ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/HHCov_Statewise_Tranche `tranche'_`c(current_date)'.csv", ///
cells("count mean(fmt(%13.2fc)) sd(fmt(%13.2fc))") nonumber ///
nomtitle nonote noobs label collabels("N" "Mean" "SD") replace

restore


********************************************************************************
* SECTION 4 - COMBINE TABLES INTO ONE EXCEL SHEET
********************************************************************************
global tables MH_RawScaleScore MH_ElevScale TimeUse_all Health_Status_all ADLs_all ///
Immu_ANC_all Immu_ANC_bygender IndCov_all HHCov_Simple_Summ HHCov_Statewise

foreach table of global tables {
import delimited ///
"$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Raw tables/`table'_Tranche `tranche'_`c(current_date)'.csv", clear
	* remove quotes from each cell
	foreach var of varlist * {
		replace `var' = substr(`var', 3,.)
		replace `var' = substr(`var', 1,strlen(`var') - 1)
	}
export excel using "$sehat/SEHAT Study/Output/Wave 3 - Oct 2022/Preliminary Tables_Tranche `tranche'_`c(current_date)'.xlsx", sheet("`table'") 
}






log close