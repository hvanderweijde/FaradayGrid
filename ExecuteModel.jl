##############################################################################
###                            Faraday Grid                                ###
###                       Wholesale market model                           ###
###                            Control panel                               ###
###                                v0.1                                    ###
##############################################################################

##############################################################################
###                               options                                  ###
##############################################################################
input_data1 = "/Users/harry/Dropbox/Faraday Grid/Wholesale market model/InputDataUK3.xlsx" #input data file
input_data2 = "/Users/harry/Dropbox/Faraday Grid/Wholesale market model/InputDataUK3.xlsx" #input data file
input_demandfunctions="/Users/harry/Dropbox/Faraday Grid/Wholesale market model/DemandFunctions.xlsx"
input_marketoutcomes="/Users/harry/Dropbox/Faraday Grid/Wholesale market model/MarketOutcome.xlsx"

output_marketoutcomes= "/Users/harry/Dropbox/Faraday Grid/Wholesale market model/MarketOutcome.xlsx" #output data file
output_redispatch= "/Users/harry/Dropbox/Faraday Grid/Wholesale market model/RedispatchOutcome.xlsx"
output_demandfunctions="/Users/harry/Dropbox/Faraday Grid/Wholesale market model/DemandFunctions.xlsx" #output data file

epsilon = -0.2 #demand elasticity
hh=24 #length of horizon in market model in hours, must be between 1 and 8761 but higher numbers may lead to very long solution times
rhh=1 #length of horizon in redispatch model in hours, must be between 1 and 8761 but higher numbers may lead to very long solution times
elastic="false"

##############################################################################
###                               run models                               ###
##############################################################################


include("ModelFunctions.jl")

# Data input
ProcessData(input_data1) #process input data - necessary before all model runs
LoadDemandFunctions(input_demandfunctions) #load previously saved demand functions - only necessary if elastic demand model is used
LoadMarketOutcomes(input_marketoutcomes) #load previously saved market outcomes - only necessary if redispatch model is used without market model

# Data output
WriteMarketData(output_marketoutcomes) #write market outcomes to specified file
WriteRedispatchData(output_redispatch) #write redispatch decisions to specified file
WriteDemandFunctions(output_demandfunctions,epsilon) #generate and write estimated demand functions to specified file

# Plot outputs
GeneratePricePlot() #generate plot of prices over time
GenerateRedispatchCostPlot() #generate plot of redispatch costs over time

# Models
RollingHorizonModel(hh) #run rolling horizon model with specified horizon length
RollingHorizonModelElasticDemand(hh) #run rolling horizon model with elastic demand with specified horizon length
ImperfectCompetitionModel(hh) #run Nash-Cournot imperfectly competitive market model
RedispatchModel(rhh,elastic) #run redispatch model with specified horizon length
