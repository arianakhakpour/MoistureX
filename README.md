# About MoistureX
The objective of this project is to develop a computer model to calculate the relative humidity within an LED enclosure for various climate conditions. Condensed water vapour within an LED unit may lead to permanent damage to the units printed circuit board. The developed model determines the internal relative humidity of the LED enclosure based on the climate data file and the LED units operative schedule.

A vent is located on the back of the LED enclosure to allow pressure relief. This membrane is permeable and allows the passage of water vapour through it. Once the local relative humidity within a unit reaches 100%, condensation will occur. The amount of vapour permeating through the vents is dependent on the vent’s physical properties. An experimental investigation was performed to characterize the vapour permeability coefficient for the vent at the University of Toronto. These results were then applied to develop an equation to calculate the flow rate of moisture through the permeable membrane, based on the ambient and internal conditions. 

Next, the LED unit was tested in an environmental chamber at GVA. For a given ambient relative humidity and temperature, both the vent’s permeability and the internal temperature of the chamber are required to accurately calculate the relative humidity within the enclosure. Using the environmental chamber at GVA, the internal temperature and relative humidity of the LED were recorded to study moisture transfer for various climate and operative settings. Environmental tests at a constant humidity were conducted to study the thermal profile of the LEDs at different ambient temperature during operation. This knowledge further enhanced the accuracy of the developed model.

The developed model was validated using experimental data, an average error of less than 20% for relative humidity and 3% for temperature were noted. The model is computationally efficient and can be simply adapted for different vents or LED units. The model assumes a clean vent and stagnant ambient airflow with no external convection. During natural operation, the vent can become partially clogged due to pollution, contamination and droplet nucleation within their structure, changing its permeability. Therefore, simulation uncertainty may increase during the vent lifecycle.

# Software Implementation
A computer program was developed (using Python 3.7) to simulate the LED’s internal relative humidity profile for a given climate data file and a user-specified daily operation schedule. The code utilizes an iterative approach to the internal humidity modelling technique described in the previous section. 


# Inputs and Outputs
To run the computer simulation, user inputs are required. User inputs for this computer program are the annual climate data file, Gore-Vent’s active area, and the LED’s switch-on and switch-off times. After running the code, an Excel file is generated that contains the forecasted internal relative humidity values for every hour of the year. The secondary output that will be displayed on the Python compiler is a list of hours that convey a higher risk of internal condensation (RH > 90%).

# Please view "MoistureX software instruction" file for step by step instructions and more details.

