SELECT PartDepListing.*, MHPPartCov.PEFROMDT, MHPPartCov.PETODATE, MHPPartCov.PECOVCAT1, MHPPartCov.PECOVCAT2, MHPPartCov.PECOVCAT3, MHPPartCov.PECOVCAT4, MHPPartCov.PECOVCAT5, 
MHPPartCov.PECOVCAT6, MHPPartCov.PECOVCAT7, MHPPartCov.PECOVCAT8, MHPPartCov.PECOVCAT9, MHPPartCov.PECOVCAT10, MHPPartCov.PECOVCAT11, MHPPartCov.PECOVCAT12, MHPPartCov.PECOVCAT13, 
MHPPartCov.PEENRLEV1, MHPPartCov.PEENRLEV2, MHPPartCov.PEENRLEV3, MHPPartCov.PEENRLEV4, MHPPartCov.PEENRLEV5, MHPPartCov.PEENRLEV6, MHPPartCov.PEENRLEV7, MHPPartCov.PEENRLEV8, 
MHPPartCov.PEENRLEV9, MHPPartCov.PEENRLEV10, MHPPartCov.PEENRLEV11, MHPPartCov.PEENRLEV12, MHPPartCov.PEENRLEV13 
INTO MHPPartDepAndCov
FROM PartDepListing INNER JOIN MHPPartCov ON (PartDepListing.GPGROUP = MHPPartCov.GPGROUP) AND (PartDepListing.GPPARTIC = MHPPartCov.GPPARTIC)
WHERE (((PartDepListing.GPBENEND)="") AND ((PartDepListing.DPTRMDTE)=""));
