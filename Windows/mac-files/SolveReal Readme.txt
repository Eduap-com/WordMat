solvereal package by Mikael Samsøe Sørensen. www.eduap.com
GNU General Public License
Basically af package developed to improve equation-solving for real numbers.
It was originally made as part of the WordMat project (www.eduap.com)


variables:
autonsolve
	default is false. Determines whether solvereal should apply nsolve when unable to solve symbolically.

AllTrig
	default is false. Determines whether all or only 1 solution should be given to trigonometric equations

showassumptions
	default is true. Shows any assumptions for the solutions

solverealmaximasolve
	default true - solvereal uses Maxima's solve function as part of the attempt to solve the equation

solverealtopoly
	default true - Wheather solvereal should Use to_poly_solve as part of the attemps to solve the equation

Functions:
Solve(equation,var)
	if domain=complex it calls solve(equation,var)
	if domain=real it calls solvereal(equation,var)
	Also Works for system of equations

solvereal(equation,var)
	Solves equation for var, wihtin real numbers.
	if given list of equations and variables solvesystem is automatically applied

nsolve(equation,var)		
	Solves the equation by computing large number of values.Then applying realroots

nsolve(equation,var,n,m,time,maxsol,nn,mm)	
	n,m is is the min/max. order to search 10^-n - 10^m thouroughly, default n=15
	time is the max time to search, but this doesn't mean it cant halt if extremely large/small values are produced. 	maxsol is the maximum number of solutions to solve for.
	nn and mm are the max and min orders were solutions are sought loosely

solvesystem(equations,vars)
	Solves systems of equations. example: solvesystem([x^2+y^2=1,y=x],[x,y])
	

simplify(expression)		Simplfies using different methods. Returning the shortest/most compact answer.

desolvemore()			solves 1. and 2. order DE
				uses combination of ode2, solvereal, contrib_ode, ic1,ic2,bc2
				Examples of syntax:
				desolvemore('diff(y,x)=y,x,y);
				desolvemore('diff(y,x)=y,x=0,y=1);		 1. order initial value
				desolvemore('diff(y,x,2)=y,x=0,y=1,x=2,y=4);	 2. order initial value
				desolvemore('diff(y,x,2)=y,x=0,y=1,diff(y,x)=2); 2. order boundary value