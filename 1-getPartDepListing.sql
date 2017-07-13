SELECT PARTICIP.GPGROUP, PARTICIP.GPPARTIC, PARTICIP.GPLAST, PARTICIP.GPFIRST, PARTICIP.GPMI, PARTICIP.GPSSN, PARTICIP.GPBENEFF, PARTICIP.GPBENEND, DEPENDNT.DPDEP, 
DEPENDNT.DPSSN, DEPENDNT.DPLAST, DEPENDNT.DPFIRST, DEPENDNT.DPMI, DEPENDNT.DPDOB, DEPENDNT.DPSEX, DEPENDNT.DPRELAT, DEPENDNT.DPINITDT, DEPENDNT.DPTRMDTE 
INTO PartDepListing
FROM PARTICIP INNER JOIN DEPENDNT ON (PARTICIP.GPPARTIC = DEPENDNT.DPPARTIC) AND (PARTICIP.GPGROUP = DEPENDNT.DPGROUP)
WHERE (((PARTICIP.GPGROUP)=[Group ID]) AND ((DEPENDNT.DPGROUP)=[Group ID]));
