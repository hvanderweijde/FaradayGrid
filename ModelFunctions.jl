##############################################################################
###                            Faraday Grid                                ###
###                       Wholesale market model                           ###
###                        Perfect competition                             ###
###                                v0.1                                    ###
##############################################################################


##############################################################################
###                         initialise packages                            ###
##############################################################################


using JuMP                  #necessary for optimisation
using Gurobi                #solver - requires license. Other solvers can be used
using Complementarity       #required for imperfectly competitive model. Requires free PATH license
using Taro                  #required to read and write Excel files
using Plots                 #required to generate plots
using DataFrames            #required to process input data

try                         #initialises Taro if this has not been done already
    Taro.init()
catch
    print("Taro already initialised")
end

##############################################################################
###                             I/O modules                                ###
##############################################################################
function ProcessData(input_data)
    grid_data=Taro.readxl(input_data, "Network", "A2:E53"; header=false)
        global lines =             Array(grid_data[1])
        global from =              Array(grid_data[2])
        global to =                Array(grid_data[3])
        reactance_raw =     Array(grid_data[4])
        maxflow_raw =       Array(grid_data[5])
    generator_data=Taro.readxl(input_data, "Generators", "A2:H217"; header=false)
        global generators =        Array(generator_data[1])
        global owner =             Array(generator_data[2])
        global location =          Array(generator_data[3])
        fuelcost_raw=       Array(generator_data[4])
        capacity_raw=       Array(generator_data[5])
        global variableoutput_raw= Array(generator_data[6])
        uprate_raw=         Array(generator_data[7])
        downrate_raw=       Array(generator_data[8])
    global availability_raw = Array(Taro.readxl(input_data, "Availability", "B2:P8762"; header=false))
    load_data = Taro.readxl(input_data, "Load", "A2:AD8762"; header=false)
        global timeperiods =       Array(load_data[1])
        demand_raw =        Array(load_data[2:30])
    global buses = Array(Taro.readxl(input_data, "Load", "B1:AD1"; header=false))
    # transform input data into dictionaries
    global reactance =     Dict(zip(lines,reactance_raw))
    global maxflow =       Dict(zip(lines,maxflow_raw))
    global fuelcost =      Dict(zip(generators,fuelcost_raw))
    global capacity =      Dict(zip(generators,capacity_raw))
    global variableoutput= Dict(zip(generators,variableoutput_raw))
    global uprate =        Dict(zip(generators,uprate_raw))
    global downrate =      Dict(zip(generators,downrate_raw))

    global demand=Dict()
        for i in 1:length(buses), j in 1:length(timeperiods)
            demand[buses[i],timeperiods[j]]=demand_raw[j,i]
        end

    global variablegens=[]
        for g in generators
            if variableoutput[g]==1
                push!(variablegens, g)
            end
        end

    global availability = Dict()
        for g in 1:length(generators), j in 1:length(timeperiods)
                availability[generators[g],timeperiods[j]]=1
        end
        for g in 1:length(variablegens), j in 1:length(timeperiods)
                        availability[variablegens[g],timeperiods[j]]=availability_raw[j,g]
        end

    # generate additional dictionaries
    global firms =         unique(owner)                                  #generate list of unique owners using index f
    global gensowned =     Dict()                                         #returns list of generators owned by a particular owner
        for ff in 1:length(firms)
            ow1=[]
            for gg in 1:length(generators)
                if owner[gg]==firms[ff]
                    push!(ow1, generators[gg])
                end
            end
            gensowned[firms[ff]]=ow1
        end
    global ownedby =       Dict(zip(generators,owner))                     #return owner of a particular generator
    global fromd =         Dict(zip(lines,from))                           #generate dictionaries that return the origin and end buses of a given line
    global tod =           Dict(zip(lines,to))
    global connectedto =   Dict()                                          #generate dictionary that returns lists of lines connected to a given bus
        for i in 1:length(buses)                                        #outer loop over buses
            aux1=[]                                                     #temporary auxilliary array
            for l in 1:length(lines)                                    #inner loop over lines
                if to[l]==buses[i]                                      #check if line ends in bus
                    push!(aux1,lines[l])                                #if it does, add its name to the temporary auxilliary variable
                end
            end
            connectedto[buses[i]]=aux1                                  #write contents of auxilliary variable to dictionary
        end
    global connectedfrom=  Dict()                                          #generate dictionary that returns lists of lines connected from a given bus
        for i in 1:length(buses)                                        #outer loop over buses
            aux2=[]                                                     #temporary auxilliary array
            for l in 1:length(lines)                                    #inner loop over lines
                if from[l]==buses[i]                                    #check if line originates in bus
                    push!(aux2,lines[l])                                #if it does, add its name to the temporary auxilliary variable
                end
            end
            connectedfrom[buses[i]]=aux2                                #write contents of auxilliary variable to dictionary
        end
    global locatedat=Dict()                                                #generate dictionary that returns list of generators for a given bus
        for i in 1:length(buses)                                        #outer loop over buses
            aux1=[]                                                     #temporary auxilliary array
            for g in 1:length(generators)                               #inner loop over generators
                if location[g]==buses[i]                                #check if generator is at in bus
                    push!(aux1,generators[g])                           #if it is, add its name to the temporary auxilliary variable
                end
            end
            locatedat[buses[i]]=aux1                                    #write contents of auxilliary variable to dictionary
        end
    global qpeak=Dict()
        for i in buses
            qpeak[i]=maximum([demand[i,t] for t in timeperiods])
        end
end

function LoadDemandFunctions(demand_functions)
    consObjA_raw = Array(Taro.readxl(demand_functions, "Intercepts", "C3:AE8763"; header=false))
    consObjB_raw = Array(Taro.readxl(demand_functions, "Slopes", "C3:AE8763"; header=false))
    global consObjA=Dict()
        for i in 1:length(buses), j in 1:length(timeperiods)
            consObjA[buses[i],timeperiods[j]]=consObjA_raw[j,i]
        end
    global consObjB=Dict()
        for i in 1:length(buses), j in 1:length(timeperiods)
            consObjB[buses[i],timeperiods[j]]=consObjB_raw[j,i]
        end
end

function LoadMarketOutcomes(output_file)
    marketdata_raw = Array(Taro.readxl(output_file, "MarketOutcome", "C3:HJ8763"; header=false))
    global gref=Dict()
        for g in 1:length(generators), j in 1:length(timeperiods)
            gref[generators[g],timeperiods[j]]=marketdata_raw[j,g]
        end
end

function WriteMarketData(output_data_market)
    w=Workbook()
    s=createSheet(w, "MarketOutcome")
    r=createRow(s, 1)
    for g in 1:length(generators)
        c=createCell(r, g+1); setCellValue(c, generators[g])
    end
    for t in 1:length(timeperiods)
        r=createRow(s, t+1)
        c=createCell(r, 1); setCellValue(c, timeperiods[t])
        for g in 1:length(generators)
            c=createCell(r, g+1); setCellValue(c, gref[generators[g],timeperiods[t]])
        end
    end
    s=createSheet(w,"Prices")
    for t in 1:length(timeperiods)
        r=createRow(s,t)
        c=createCell(r, 1); setCellValue(c, timeperiods[t])
        c=createCell(r, 2); setCellValue(c, pref[timeperiods[t]])
    end
    write(output_data_market, w)
end

function WriteRedispatchData(output_data_redispatch)
    w=Workbook()
    s=createSheet(w, "Redispatch")
    r=createRow(s, 1)
    for g in 1:length(generators)
        c=createCell(r, g+1); setCellValue(c, generators[g])
    end

    for t in 1:length(timeperiods)
        r=createRow(s, t+1)
        c=createCell(r, 1); setCellValue(c, timeperiods[t])
        for g in 1:length(generators)
            c=createCell(r, g+1); setCellValue(c, qredi[generators[g],timeperiods[t]])
        end
    end
    s=createSheet(w, "RedispatchCosts")
    for t in 1:length(timeperiods)
        r=createRow(s, t)
        c=createCell(r, 1); setCellValue(c, timeperiods[t])
        c=createCell(r, 2); setCellValue(c, credi[timeperiods[t]])
    end
    write(output_data_redispatch, w)
end

##############################################################################
###                rolling horizon wholesale market model                  ###
##############################################################################
function RollingHorizonModel(hh)
    zz=1                                                #iterator
    global gref=Dict()
    global pref=Dict()

    while zz <= length(timeperiods)
        if zz+hh-1<=length(timeperiods)
            timeperiodssubset=timeperiods[zz:zz+hh-1]
        else
            timeperiodssubset=timeperiods[zz:length(timeperiods)]
        end
        m1 = Model(solver=GurobiSolver())
        @variable(m1,     generation[g in generators, t in timeperiodssubset]  >=0)
        @objective(m1,    Min, sum(generation[g,t]*fuelcost[g] for g in generators, t in timeperiodssubset))
        @constraint(m1,   maxgen[g in generators, t in timeperiodssubset],    capacity[g]*availability[g,t]-generation[g,t]>=0)
        @constraint(m1,   marketclearing[t in timeperiodssubset],      sum(generation[g,t] for g in generators)-sum(demand[i,t] for i in buses) == 0 )
        solve(m1)
        for xx in 1:length(generators)
             gref[generators[xx],timeperiods[zz]] = getvalue(generation)[generators[xx],timeperiods[zz]]
        end
        pref[timeperiods[zz]] = getdual(marketclearing)[timeperiods[zz]]
        zz += 1
    end
end

##############################################################################
###          calculate demand functions  and write to Excel                ###
##############################################################################

function WriteDemandFunctions(output_data_demandfunctions,epsilon)

    consObjA = Dict() #calculate demand function intercept
    consObjB = Dict() #calculate demand function slope
        for i in buses, t in timeperiods[1:length(timeperiods)]
            consObjB[i,t] = (-1) * pref[t] / (demand[i,t] * epsilon)
            consObjA[i,t] = pref[t] + consObjB[i,t] * demand[i,t]
        end

    w=Workbook()
    s=createSheet(w, "Intercepts")
    r=createRow(s, 1)
    for i in 1:length(buses)
        c=createCell(r, i+1); setCellValue(c, buses[i])
    end
    for t in 1:length(timeperiods)
        r=createRow(s, t+1)
        c=createCell(r, 1); setCellValue(c, timeperiods[t])
        for i in 1:length(buses)
            c=createCell(r, i+1); setCellValue(c, consObjA[buses[i],timeperiods[t]])
        end
    end
    s=createSheet(w, "Slopes")
    r=createRow(s, 1)
    for i in 1:length(buses)
        c=createCell(r, i+1); setCellValue(c, buses[i])
    end
    for t in 1:length(timeperiods)
        r=createRow(s, t+1)
        c=createCell(r, 1); setCellValue(c, timeperiods[t])
        for i in 1:length(buses)
            c=createCell(r, i+1); setCellValue(c, consObjB[buses[i],timeperiods[t]])
        end
    end
    write(output_data_demandfunctions, w)
end



##############################################################################
###       rolling horizon wholesale market model - elastic demand          ###
##############################################################################

function RollingHorizonModelElasticDemand(hh)

    zz=1                                                #iterator
    global gref=Dict()
    global qref=Dict()
    global pref=Dict()

    while zz<=length(timeperiods)
        if zz+hh-1<=length(timeperiods)
            timeperiodssubset=timeperiods[zz:zz+hh-1]
        else
            timeperiodssubset=timeperiods[zz:length(timeperiods)]
        end

        m3 = Model(solver=GurobiSolver())

        @variable(m3,     generation[g in generators, t in timeperiodssubset]  >=0)
        @variable(m3,     rdemand[i in buses, t in timeperiodssubset]  >=0)

        @objective(m3,    Max, sum(consObjA[i,t]*rdemand[i,t]-0.5*consObjB[i,t]*rdemand[i,t]*rdemand[i,t] for i in buses, t in timeperiodssubset)-sum(generation[g,t]*fuelcost[g] for g in generators, t in timeperiodssubset))
        @constraint(m3,   maxgen[g in generators, t in timeperiodssubset],    capacity[g]*availability[g,t]-generation[g,t]>=0)
        @constraint(m3,   marketclearing[t in timeperiodssubset],      sum(generation[g,t] for g in generators)-sum(rdemand[i,t] for i in buses) == 0 )
        solve(m3)

        for xx in 1:length(generators)
             gref[generators[xx],timeperiods[zz]] = getvalue(generation)[generators[xx],timeperiods[zz]]
        end

        for ii in 1:length(buses)
            qref[buses[ii],timeperiods[zz]]=getvalue(rdemand)[ii,zz]
        end

        pref[timeperiods[zz]] = getdual(marketclearing)[timeperiods[zz]]

        zz+=1
    end


end



##############################################################################
###                          redispatch model                              ###
##############################################################################

function RedispatchModel(rhh,elastic)

    if elastic=="false"
        demand2=demand
    end

    if elastic=="true"
        demand2=qref
    end


    zz=1                                                #iterator
    global qredi=Dict()
    global credi=Dict()

    while zz <= length(timeperiods)
        if zz+rhh-1<=length(timeperiods)
            timeperiodssubset=timeperiods[zz:zz+rhh-1]
        else
            timeperiodssubset=timeperiods[zz:length(timeperiods)]
        end
        m2 = Model(solver=GurobiSolver())
        @variable(m2,    voltageangle[i in buses, t in timeperiodssubset])
        @variable(m2,    redispatch[g in generators, t in timeperiodssubset])
        @expression(m2,  flow[l in lines,t in timeperiodssubset],         (1/reactance[l])*(voltageangle[fromd[l],t]-voltageangle[tod[l],t]))
        @objective(m2, Min, sum(fuelcost[g]*redispatch[g,t] for g in generators, t in timeperiodssubset))
        @constraint(m2,  slackbus[t in timeperiodssubset],                voltageangle[buses[1],t] ==  0)
        @constraint(m2,  nodalbalance[i in buses, t in timeperiodssubset],demand2[i,t]-sum(gref[g,t]+redispatch[g,t] for g in locatedat[i]) + reduce(+,0,flow[l,t] for l in connectedto[i])-reduce(+,0,flow[l,t] for l in connectedfrom[i]) ==  0)
        @constraint(m2,  maxflowup[l in lines, t in timeperiodssubset],   flow[l,t] <=  maxflow[l])
        @constraint(m2,  maxflowdown[l in lines, t in timeperiodssubset], flow[l,t] >= -maxflow[l])
        @constraint(m2,  maxgen[g in generators, t in timeperiodssubset], gref[g,t]+redispatch[g,t] <=  capacity[g]*availability[g,t])
        @constraint(m2,  mingen[g in generators, t in timeperiodssubset], gref[g,t]+redispatch[g,t] >=  0)
        solve(m2)
        for xx in 1:length(generators)
             qredi[generators[xx],timeperiods[zz]] = getvalue(redispatch)[generators[xx],timeperiods[zz]]
        end
        credi[timeperiods[zz]]=getobjectivevalue(m2)
        zz+=1
    end
end


##############################################################################
###                    imperfectly competitive model                       ###
##############################################################################

function ImperfectCompetitionModel(hh)
    zz=1                                                #iterator
    global gref=Dict()
    global pref=Dict()
    global qref=Dict()

    while zz <= length(timeperiods)
        if zz+hh-1<=length(timeperiods)
            timeperiodssubset=timeperiods[zz:zz+hh-1]
        else
            timeperiodssubset=timeperiods[zz:length(timeperiods)]
        end

        m4=MCPModel()

        @variable(m4,   generation[g in generators, t in timeperiodssubset]   >=  0)
        @variable(m4,   lambda[g in generators, t in timeperiodssubset]       >=  0)

        @mapping(m4,    kktgen[g in generators,t in timeperiodssubset],      consObjA[buses[1],t]+(1/sum(1/consObjB[i,t] for i in buses))*(sum(generation[gg,t] for gg in gensowned[ownedby[g]])+sum(generation[gg,t] for gg in generators))+fuelcost[g]+lambda[g,t])
        @mapping(m4,    kktlambda[g in generators,t in timeperiodssubset],   capacity[g]*availability[g,t]-generation[g,t])

        @complementarity(m4,    kktgen,         generation)
        @complementarity(m4,    kktlambda,      lambda)

        PATHSolver.options(convergence_tolerance=1e-8, output="no", time_limit=3600)

        solveMCP(m4)

        for xx in 1:length(generators)
             gref[generators[xx],timeperiods[zz]] = getvalue(generation)[generators[xx],timeperiods[zz]]
        end
        pref[timeperiods[zz]] = consObjA[buses[1],timeperiods[zz]]+(1/sum(1/consObjB[i,timeperiods[zz]] for i in buses))*sum(gref[g,timeperiods[zz]] for g in generators)
        for ii in 1:length(buses)
            qref[buses[ii],timeperiods[zz]] = consObjA[buses[ii],timeperiods[zz]]+consObjB[buses[ii],timeperiods[zz]]*pref[timeperiods[zz]]
        end


    end
end





##############################################################################
###                            generate plots                              ###
##############################################################################
function GeneratePricePlot()
    plot([pref[t] for t in timeperiods[1:length(timeperiods)]],title="Price",leg=false)
    savefig("priceplot.png")
end

function GenerateRedispatchCostPlot()
    plot([credi[t] for t in timeperiods[1:length(timeperiods)]],title="Redispatch costs",leg=false)
    savefig("redispatchcostplot.png")
end

function GenerateDemandPlot()
    plot([sum(demand[i,t] for i in buses) for t in timeperiods],title="Demand",leg=false)
    savefig("demandplot.png")
end
