'   Copyright 2014
'
'   Gregory Hicks
' Software Engineer
'
' Licensed under Creative Commons
'
' You may use or extend this work under the following conditions:
' 1. You must include this copyright in your derived works.
' 2. The software is made available "As-Is". 
'    No liabilities accepted for your use or updates applied
'    No Warranties, implicit or implied apply to use.
'
' Cucumber Scenarios
' Feature: initialise DLL
' Feature: Open QC Connection
' Feature: Open Domain and Project
' Feature: Locate The Test set
' Feature: NextCase
' Feature: NextStep
' Feature: EOTS
' Feature: EOTC

Function newADO
	set newADO = new clsADO
end function

class clsADO

	private objConnection , bEOTC , bEOTS
	
	sub class_initialize
		SET objConnection = CreateObject("TDApiOle80.TDConnection")
	end sub
	
	' used to update the registry with a hack to enable using DLL surrogacy
	sub InitialiseRegistry
		SET objConnection = CreateObject("TDApiOle80.TDConnection")
	end sub
	
	sub class_terminate : catch
	end sub
	
	sub catch
		if err.number = 0 then exit sub
	end sub
	
	'@ param
	'@ param
	'@ param
	sub OTAConnect ( sURL , sUID , sPWD )
		
	end sub
	
	'@ param
	'@ param
	'@ param
	sub OpenDomainProject ( sDomain , sProject )
		
	end sub
	
	'@ param
	'@ param
	'@ param
	sub locateTestSet ( sTestSetname )
	
	end sub
	
	sub NextCase
	
	end sub
	
	property get  EOTC
	
	end sub
	
	sub NextStep
		
	end sub
	
	property get EOTS
	
	end sub
	
	
end class
