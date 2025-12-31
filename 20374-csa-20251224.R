#######################################################  VARIABLES ##################################################################
El.Sub.ID <- 20374
# update working directory
WD <- paste0(
  Working.Folder,"/",El.Sub.ID)
setwd(WD)
#######################################################  PACKAGES ################################################################
library(tidyverse)
library(sf)
library(openxlsx)
library(terra)
library(DBI)
library(odbc)
library(glue)
library(lwgeom)
##################################################  RANGE EXTENT ###############################################################
#select layer
BEC13<- st_read(dsn = BEC13.path, layer = "BEC13_v2")
#Queries for Range extent
BGC.Units.Min <- c("ESSFdc1", "ESSFdc2", "ESSFmh", "ESSFwm2", "ESSFxc2", "ICHdw4", "ICHmk1", "ICHmk2", "ICHmk3", "ICHmw3", "ICHmw5", "ICHvc", "ICHwk1", "ICHxm1", "IDFdk1", "IDFdk2", "IDFdk5", "IDFdm1", "IDFdm2", "IDFxk", "MSdk", "MSdm1", "MSdm2", "MSdm3", "MSdw", "MSxk", "MSxk1", "MSxk2", "SBPSdc", "SBPSmk", "SBPSxc", "SBSdk", "SBSdw2", "SBSdw3", "SBSmc2", "SBSmk1", "SBSmw", "SBSwk1")
# minimal evidence supports occurrence of Ws04 in the BWBS. LMH 52 does not list Ws04 in the BWBS or SWB.  LMH68 does not list Ws04 or related Fl05 in any BWBS.  BWBSmw and wk1 seem to be at the NE margin of S. drummondiana and Spiraea douglasii range. Two plots support the occurrence of the Ws04 in the SBSdh1 (Kyla Rushton personal communication)
BGC.Units.Max <- c("BWBSwk1", "BWBSmw", "ESSFdc1", "ESSFdc2", "ESSFmh", "ESSFwm2", "ESSFxc2", "ICHdw4", "ICHmk1", "ICHmk2", "ICHmk3", "ICHmw3", "ICHmw5", "ICHvc", "ICHwk1", "ICHxm1", "IDFdk1", "IDFdk2", "IDFdk5", "IDFdm1", "IDFdm2", "IDFxk", "MSdk", "MSdm1", "MSdm2", "MSdm3", "MSdw", "MSxk", "MSxk1", "MSxk2", "SBPSdc", "SBPSmk", "SBPSxc","SBSdh1", "SBSdk", "SBSdw2", "SBSdw3", "SBSmc2", "SBSmk1", "SBSmw", "SBSwk1", "SBSwk2")

BGC.Range.Min <- BEC13 %>% filter(BEC13$MAP_LABEL %in% BGC.Units.Min)
BGC.Range.Max <- BEC13 %>% filter(BEC13$MAP_LABEL %in% BGC.Units.Max)
#MCP around selected BGC.Units
BGC.MCP.Min <- st_convex_hull(st_union(BGC.Range.Min))
BGC.MCP.Max <- st_convex_hull(st_union(BGC.Range.Max))
#area calculations
BGC.MCP.Min.area.km2 <- setNames(as.numeric(st_area(BGC.MCP.Min)) / 1e6, "BGC.MCP.Min.area.km2")
BGC.MCP.Max.area.km2 <- setNames(as.numeric(st_area(BGC.MCP.Max)) / 1e6, "BGC.MCP..Max.area.km2")
BGC.Range.Min.km2 <- setNames(sum(BGC.Range.Min$Shape_Area, na.rm = TRUE)/ 1e6, "BGC.Range.Min.km2")
BGC.Range.Max.km2 <- setNames(sum(BGC.Range.Max$Shape_Area, na.rm = TRUE)/ 1e6, "BGC.Range.Max.km2")
################################################## PROJECT BOUNDARIES - Ecosystem Mapping ##########
### Create SF of ALL Ecosystem Mapping Project Boundaries that overlap the widest possible BGC Range
TEI.proj.bound.all <- st_read(
  dsn   = TEI.proj.bound.path,
  query = "
    SELECT *
    FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP
    WHERE PROJECT_TYPE IN (
     'EST', 'EXC', 'NEM', 'NEMNSS', 'NEMSEI', 'NEMWHR',
      'PEM', 'PEMSDM', 'PEMTBT', 'PEMWHR', 'PREM', 'SEI', 'SEIWHR', 'STS',
      'TEM', 'TEMNSS', 'TEMSDM', 'TEMSEI', 'TEMSET', 'TEMTSM', 'TEMWHR',
      'VEG', 'WET'
    )
  "
)
# filter to those that intersect the BGC range
TEI.proj.bound.all <- st_filter(TEI.proj.bound.all, BGC.Range.Max)

###Maxiumum TEI Project Boundary/BGC Range Overlap Estimate
# This excludes PEM and other projects that inventory at a thematic scalle too coarse to capture the 20374
# It includes all TEI datasets that could conceivably detect the 20374, though it is uncertain whether all
# of these projects effectively inventory for the 20374. This only includes the core BGC Range, not the
# BGCs where occurrence of the 20374 is yet to be confirmed with confidence.
TEI.proj.bound.max <- st_read(
  dsn   = TEI.proj.bound.path,
  query = "
    SELECT *
    FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP
    WHERE BUSINESS_AREA_PROJECT_ID IN (1, 2, 4, 9, 27, 28, 79, 91, 108, 114, 115, 119, 135, 145, 149, 183, 184, 192, 196, 198, 208, 209, 213, 214, 216, 218, 219, 221, 222, 226, 231, 233, 239, 240, 244, 247, 248, 1027, 1037, 1046, 1048, 1049, 1054, 1055, 1058, 1059, 1060, 1068, 1070, 4479, 4482, 4498, 4508, 4523, 4917, 5679, 5680, 6130, 6131, 6473, 6484, 6526, 6536, 6537, 6605, 6624)
  ")
overlap_sf <- st_intersection(BGC.Range.Min, TEI.proj.bound.max)
# Dissolve overlapping geometries
overlap_sf <- st_union(overlap_sf)
# Calculate area in m²
overlap_sf$overlap_area_m2 <- st_area(overlap_sf)
# Calculate area of BGC range that is overlapped by selected project boundaries
BGC.Range.Mapped.Max.km2 <- setNames(as.numeric(sum(overlap_sf$overlap_area_m2)/ 1e6),"BGC.Range.Mapped.Max.km2")
BGC.Range.Mapped.Max.Percent <- setNames(as.numeric(BGC.Range.Mapped.Max.km2/BGC.Range.Max.km2)*100,"BGC.Range.Mapped.Max.Percent")
###Minimum TEI Project Boundary Overlap Estimate
# Create SF of selected projects
#TEI.proj.bound.selected <- st_read(dsn   = TEI.proj.bound.path,query = " SELECT * FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP WHERE BUSINESS_AREA_PROJECT_ID IN (4, 135, 216, 233, 244, 1046, 1048, 1049, 1054, 4523, 4917, 6468, 6536, 6624)")
TEI.proj.bound.min <- st_read(
  dsn   = TEI.proj.bound.path,
  query = "
    SELECT *
    FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP
    WHERE BUSINESS_AREA_PROJECT_ID IN (4, 108, 135, 209, 216, 233, 239, 244, 1046, 1048, 1049, 1054, 1055, 4523, 6473, 6536, 6605, 6624)
  ")
### determine MIN extent of overlap between selected ecosystem mapping projects and BGC Range 
overlap_sf <- st_intersection(BGC.Range.Min, TEI.proj.bound.min)
# Dissolve overlapping geometries
overlap_sf <- st_union(overlap_sf)
# Calculate area in m²
overlap_sf$overlap_area_m2 <- st_area(overlap_sf)
# Calculate area of BGC range that is overlapped by selected project boundaries
BGC.Range.Mapped.Min.km2 <- setNames(as.numeric(sum(overlap_sf$overlap_area_m2)/ 1e6),"BGC.Range.Mapped.Min.km2")
BGC.Range.Mapped.Min.Percent <- setNames(as.numeric(BGC.Range.Mapped.Min.km2/BGC.Range.Min.km2)*100,"Percent.BGC.Range.Mapped.Min")

################################################  BECMASTER - QUERY AND CLEAN ##################################################
# Connect to BECMaster
BEC.Master.Connect <- dbConnect(odbc::odbc(),
                                   .connection_string = paste0("Driver={Microsoft Access Driver (*.mdb, *.accdb)};",
                                                               "DBQ=", BECMaster.path, ";"))
# Query ENV for focal units. Note that this query does not include some plots identified by BECSiteUnit in the Admin table
# valid plots from the Admin table which were not returned by this query were manually added by plot number in the next query
BEC.Master.Focal.Query1 <- "
SELECT 
  e.PlotNumber, e.Date, e.SiteSurveyor, e.Zone, e.SubZone, e.SiteSeries, e.PlotRepresenting, e.Location, 
  e.Longitude, e.Latitude, e.UTMZone, e.UTMEasting, e.UTMNorthing, e.LocationAccuracy, 
  e.SiteModifier1, e.SiteModifier2, e.MapUnit, e.MoistureRegime, e.NutrientRegime, 
  e.StructuralStage, e.StructuralStageMod, e.StandAge, e.Elevation, e.SlopeGradient, e.Aspect, 
  e.MesoSlopePosition, e.Exposure1, e.Exposure2, e.SiteDisturbance1, e.SiteDisturbance2, 
  e.SiteDisturbance3, e.SiteNotes, e.VegSurveyor, e.StrataCoverTree, e.StrataCoverShrub, 
  e.StrataCoverHerb, e.StrataCoverMoss, e.VegNotes, a.BECSiteUnit, a.UserSiteUnit

FROM BECMaster19_Env AS e
LEFT JOIN BECMaster19_Admin AS a
  ON e.PlotNumber = a.Plot
WHERE e.SiteSeries LIKE '%Ws04%'
"
BEC.Master.Focal1 <- dbGetQuery(BEC.Master.Connect, BEC.Master.Focal.Query1)
BEC.Master.Focal.Query2 <- "
SELECT 
  e.PlotNumber, e.Date, e.SiteSurveyor, e.Zone, e.SubZone, e.SiteSeries, e.PlotRepresenting, e.Location, 
  e.Longitude, e.Latitude, e.UTMZone, e.UTMEasting, e.UTMNorthing, e.LocationAccuracy, 
  e.SiteModifier1, e.SiteModifier2, e.MapUnit, e.MoistureRegime, e.NutrientRegime, 
  e.StructuralStage, e.StructuralStageMod, e.StandAge, e.Elevation, e.SlopeGradient, e.Aspect, 
  e.MesoSlopePosition, e.Exposure1, e.Exposure2, e.SiteDisturbance1, e.SiteDisturbance2, 
  e.SiteDisturbance3, e.SiteNotes, e.VegSurveyor, e.StrataCoverTree, e.StrataCoverShrub, 
  e.StrataCoverHerb, e.StrataCoverMoss, e.VegNotes, a.BECSiteUnit, a.UserSiteUnit

FROM BECMaster19_Env AS e
LEFT JOIN BECMaster19_Admin AS a
  ON e.PlotNumber = a.Plot
WHERE a.BECSiteUnit LIKE '%Ws04%' OR a.UserSiteUnit LIKE '%Ws04%'
"
BEC.Master.Focal2 <- dbGetQuery(BEC.Master.Connect, BEC.Master.Focal.Query2)
# bind BECmaster queries from ENV and ADMIN tables, keep only unique records.
BEC.Master.Focal <-unique(rbind(BEC.Master.Focal1,BEC.Master.Focal2))
# Force longitudes to be negative, lat to be numeric
BEC.Master.Focal$Longitude <- -abs(as.numeric(BEC.Master.Focal$Longitude))
BEC.Master.Focal$Latitude  <- as.numeric(BEC.Master.Focal$Latitude)
# REMOVES ROWS WITH NO COORDS. Convert to sf object using Longitude and Latitude
BEC.Master.Focal.sf <- BEC.Master.Focal %>%
  filter(is.finite(Longitude), is.finite(Latitude)) %>%
  st_as_sf(coords = c("Longitude", "Latitude"), crs = 4326)
# Reproject to albers
BEC.Master.Focal.sf <- st_transform(BEC.Master.Focal.sf, crs = 3005)
# Buffer each point by half the LocationAccuracy for viewing in GIS
BEC.Master.Focal.sf.buffer <- BEC.Master.Focal.sf %>%
  mutate(geometry = st_buffer(geometry, dist = LocationAccuracy / 2))
# list of plots vetted either by BECSiteUnit field or by Kristi Iverson email 20251125
BEC.Vetted <- c(
  '14-3231','14-5469','272980','378480','8400061','9613288','9613857','9615182','9615223',
  '9617072','9628709','9628714','9628718','9628877','9628878','9628880','963106','9650957',
  'CL00015','CL00035','CL00041','CL00061','CL00062','CL00063','N05-182'
)
# Add new field 'Include' to filter vetted plots 
BEC.Master.Focal <- BEC.Master.Focal %>%
  mutate(
    Include = dplyr::if_else(
      PlotNumber %in% BEC.Vetted, # The logical test
      "Yes",                      # Value if TRUE
      "No"                        # Value if FALSE
    )
  )
################################################# TEI - QUERY AND CLEAN #########################################################
### TEM ###
# These queries take about 1 hour to run
# These queries warnings that don't seem to matter; "In CPL_read_ogr(dsn, layer, query, as.character(options), quiet,  :GDAL Error 1: Error occurred in filegdbtable.cpp at line 1777"
### Started with a wildcard query (~LIKE '%Ws04%') not required as only clean values (Ws04) returned 20251211
# do not need to query SITEMC fields as they return no values, but note that they do in 6473 and 6605 which are manually entered below
Query.TEM.s <- r"(
SELECT
  SDEC_1, SDEC_2, SDEC_3,
  SITE_S1, SITE_S2, SITE_S3,
  STRCT_S1, STRCT_S2, STRCT_S3,
  STAND_A1, STAND_A2, STAND_A3,
  SITEMC_S1, SITEMC_S2, SITEMC_S3,
  BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE,
  PLOT_NO, POLY_COM,
  PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,
  Shape_Area, shape
FROM TEIS_Master_Long_Tbl
WHERE PROJ_TYPE IN (
  'ESA','EST','EXC','NEM','NEMNSS','NEMSEI','NEMWHR',
  'PEM','PEMSDM','PEMTBT','PEMWHR',
  'PREM','SEI','SEIWHR','STS',
  'TEM','TEMNSS','TEMSDM','TEMSEI','TEMSET','TEMTSM','TEMWHR',
  'VEG','WET'
)
AND (SITE_S1 = 'Ws04' OR SITE_S2 = 'Ws04' OR SITE_S3 = 'Ws04')
)"


TEM.s <- st_read(
  dsn = TEI.Path,
  query = Query.TEM.s,
  quiet = TRUE
)

# Create a new column that concatenates the four BGC fields
TEM.s <- TEM.s %>%
  unite(
    col = BGC.Full,  # New column name
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE, # Columns to combine
    sep = "",
    na.rm = TRUE, #not insert NA for NULL values
    remove = FALSE
  )

### Mapcode Query
# Initial query - narrows to relevant BAPIDS
Query.TEM.m <- "
SELECT 
    SDEC_1, SDEC_2, SDEC_3,
    SITE_S1, SITE_S2, SITE_S3,
    STRCT_S1, STRCT_S2, STRCT_S3,
    STAND_A1, STAND_A2, STAND_A3,
    SITEMC_S1, SITEMC_S2, SITEMC_S3,
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE,
    PLOT_NO, POLY_COM,
    PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,
    Shape_Area, shape
FROM TEIS_Master_Long_Tbl
WHERE BAPID IN (4, 108, 135, 209, 216, 233, 239, 244, 1046, 1048, 1049, 1054, 1055)
"

TEM.m <- st_read(
  dsn   = TEI.Path,
  query = Query.TEM.m,
  quiet = TRUE
)

# concatenate BGC
TEM.m <- TEM.m %>%
  unite(
    col = BGC.Full,  # New column name
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE, # Columns to combine
    sep = "",
    na.rm = TRUE, #not insert NA for NULL values
    remove = FALSE
  )

# I checked and map codes are not appearing in site code field.
# Further narrows to relevant mapcodes
fields <- c("SITEMC_S1","SITEMC_S2","SITEMC_S3")
vals <- c("DS", "WS", "WD")

TEM.m <- TEM.m %>%
  filter(if_any(all_of(fields), ~ . %in% vals))
# Further narrows to relevant BGCs
TEM.m <- TEM.m %>%
  filter(BGC.Full %in% c(
    "ESSFdc2", "ESSFxc", "IDFdk2", "IDFdm1", "MSdm1", 
    "MSdm2", "MSxk", "SBPSmk", "SBPSxc", "SBSdk", 
    "SBSdw2", "SBSmc2", "SBSmw", "SBSwk1"
  ))

# Final query for explicit BGC/BAPID/Mapcodes
filter_condition <- 
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 4 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'IDFdk2' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'MSxk' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'MSdm2' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'ESSFxc' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'ESSFdc2' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 135 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 209 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 216 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmc2' & TEM.m$BAPID == 216 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 233 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'IDFdm1' & TEM.m$BAPID == 239 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'MSdm1' & TEM.m$BAPID == 239 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 244 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'SBSwk1' & TEM.m$BAPID == 1046 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBSdw2' & TEM.m$BAPID == 1048 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBPSmk' & TEM.m$BAPID == 1049 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBPSxc' & TEM.m$BAPID == 1054 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'MSdm1' & TEM.m$BAPID == 1055 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'IDFdm1' & TEM.m$BAPID == 1055 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS'))

TEM.m <- TEM.m %>%
  filter(!!filter_condition)

### Manually add BAPID 6605
Query.TEM.6605 <- c(
  "SELECT",
  "SDEC_1, SDEC_2, SDEC_3,",
  "SITE_S1, SITE_S2, SITE_S3,",
  "STRCT_S1, STRCT_S2, STRCT_S3,",
  "STAND_A1, STAND_A2, STAND_A3,",
  "SITEMC_S1, SITEMC_S2, SITEMC_S3, BGC_VLD,",
  "PLOT_NO, POLY_COM,",
  "PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,",
  "Shape_Area, shape",
  "FROM TEI_Long_Tbl",
  "WHERE",
  " SITE_S1 = 'Ws04' OR",
  " SITE_S2 = 'Ws04' OR",
  " SITE_S3 = 'Ws04'"
)

Query.TEM.6605 <- paste(Query.TEM.6605, collapse = " ")

TEM.6605 <- st_read(
  dsn   = TEI_BAPID6605_path,
  query = Query.TEM.6605,
  quiet = TRUE
)

TEM.6605 <- TEM.6605 %>%
  rename(BGC.Full = BGC_VLD)

### Manually add BAPID 6473
Query.TEM.6473 <- c(
  "SELECT",
  "SDEC_1, SDEC_2, SDEC_3,",
  "SITE_S1, SITE_S2, SITE_S3,",
  "STRCT_S1, STRCT_S2, STRCT_S3,",
  "STAND_A1, STAND_A2, STAND_A3,",
  "SITEMC_S1, SITEMC_S2, SITEMC_S3,",
  "BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE,",
  "PLOT_NO, POLY_COM,",
  "PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,",
  "Shape_Area, shape",
  "FROM TEI_Long_Tbl_6473",
  "WHERE",
  " SITEMC_S1 = 'Ws04' OR",
  " SITEMC_S2 = 'Ws04' OR",
  " SITEMC_S3 = 'Ws04'")

Query.TEM.6473 <- paste(Query.TEM.6473, collapse = " ")

TEM.6473 <- st_read(
  dsn   = TEI_BAPID6473_path,
  query = Query.TEM.6473,
  quiet = TRUE
)

TEM.6473 <- TEM.6473 %>%
  unite(
    col = BGC.Full,  # New column name
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE, # Columns to combine
    sep = "",
    na.rm = TRUE, #not insert NA for NULL values
    remove = FALSE
  )

# bind tem from site code queries with TEM from mapcode query
TEM <- bind_rows(TEM.s.subset, TEM.m, TEM.6473, TEM.6605)
# remove bad records based on manual review
# remove records where structual stage is not LIKE %3.  Manually reviewed these records with imagery. Most were not associated with a stream or had burned. There was also a decile error in one
bad.teis <- c(25412598, 23667093, 25412351, 25412570, 25411952, 25412356, 23666961, 23666961, 23667093, 25411952, 25412351, 25412356, 25412570, 25412598, 22905043, 23667543)
TEM <- TEM %>% filter(!TEIS_ID %in% bad.teis)

# add lat/long fields
TEM <- TEM %>%
  # Calculate the centroid for each feature
  st_centroid() %>%
  # Extract the coordinates from the centroid geometry
  st_coordinates() %>%
  # Convert the coordinate matrix (X, Y) to a data frame and bind it to TEM
  as.data.frame() %>%
  bind_cols(TEM, .) %>%
  # Rename the default 'X' and 'Y' columns
  mutate(
    Longitude = X, # The X coordinate is usually Longitude
    Latitude  = Y  # The Y coordinate is usually Latitude
  ) %>%
  # drop the temporary X and Y
  select(-X, -Y)
### Fill in missing BGC based on manual review
TEM <- TEM %>%
  mutate(
    BGC.Full = case_when(
      OBJECTID %in% c(3015989) ~ "SBSdw2",
      OBJECTID %in% c(3016307, 3016308) ~ "ICHmk3",
      TRUE ~ BGC.Full
    )
  )
######################################################## AOO Sum from TEI #######################################################
### Add AOO field to TEM to multiply deciles for focal site units.
TEM <- TEM %>%
  mutate(
    AOO = case_when(
      # Conditions based on SITE
      SITE_S1 == 'Ws04' ~ SDEC_1 * 0.1 * Shape_Area,
      SITE_S2 == 'Ws04' ~ SDEC_2 * 0.1 * Shape_Area,
      SITE_S3 == 'Ws04' ~ SDEC_3 * 0.1 * Shape_Area,
      SITEMC_S1 == 'Ws04' ~ SDEC_1 * 0.1 * Shape_Area,
      SITEMC_S2 == 'Ws04' ~ SDEC_2 * 0.1 * Shape_Area,
      SITEMC_S3 == 'Ws04' ~ SDEC_3 * 0.1 * Shape_Area,
      
      #Conditions based on BGC.Full, BAPID, and SITEMC
      
      # SBSdk / WS (BAPID 4 and 216)
      BGC.Full == 'SBSdk' & BAPID %in% c(4, 209, 216) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdk' & BAPID %in% c(4, 209, 216) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdk' & BAPID %in% c(4, 209, 216) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSmw / WD (BAPID 135, 233, 244)
      BGC.Full == 'SBSmw' & BAPID %in% c(135, 233, 244) & SITEMC_S1 == 'WD' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmw' & BAPID %in% c(135, 233, 244) & SITEMC_S2 == 'WD' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmw' & BAPID %in% c(135, 233, 244) & SITEMC_S3 == 'WD' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSmc2 / WS (BAPID 216)
      BGC.Full == 'SBSmc2' & BAPID == 216 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmc2' & BAPID == 216 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmc2' & BAPID == 216 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSwk1 / DS (BAPID 1046)
      BGC.Full == 'SBSwk1' & BAPID == 1046 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSwk1' & BAPID == 1046 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSwk1' & BAPID == 1046 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSdw2 / DS (BAPID 1048)
      BGC.Full == 'SBSdw2' & BAPID == 1048 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdw2' & BAPID == 1048 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdw2' & BAPID == 1048 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBPSmk / DS (BAPID 1049)
      BGC.Full == 'SBPSmk' & BAPID == 1049 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSmk' & BAPID == 1049 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSmk' & BAPID == 1049 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBPSxc / DS (BAPID 1054)
      BGC.Full == 'SBPSxc' & BAPID == 1054 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSxc' & BAPID == 1054 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSxc' & BAPID == 1054 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # ESSFdc2 / WS (BAPID 108)
      BGC.Full == 'ESSFdc2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFdc2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFdc2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # ESSFxc / WS (BAPID 108)
      BGC.Full == 'ESSFxc' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFxc' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFxc' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # IDFdk2 / WS (BAPID 108)
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # IDFdm1 / WS (BAPID 239, 1055)
      BGC.Full == 'IDFdm1' & BAPID %in% c(239, 1055) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdm1' & BAPID %in% c(239, 1055) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdm1' & BAPID %in% c(239, 1055) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSdm1 / WS (BAPID 239, 1055)
      BGC.Full == 'MSdm1' & BAPID %in% c(239, 1055) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm1' & BAPID %in% c(239, 1055) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm1' & BAPID %in% c(239, 1055) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSxk / WS (BAPID 108)
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSdm2 / WS (BAPID 108)
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      # Default case
      TRUE ~ 0.0 
    )
  )
# add Km2 field
TEM <-TEM %>%
  mutate(AOO.km2 = TEM$AOO/1e6)

# Sum AOO Max
AOO.Obs.Max.km2 <- sum(as.numeric(TEM$AOO),na.rm = TRUE)/1e6
# Sum AOO Min based on removal of BWBS. See comments under Range Extent
AOO.Obs.Min.km2 <- sum(as.numeric(TEM$AOO[TEM$BGC_ZONE != 'BWBS' & TEM$BGC_ZONE != 'SBSwk2']), na.rm = TRUE) / 1e6

AOO.Obs.Min.km2 <- setNames(AOO.Obs.Min.km2, "AOO.Obs.Min.km2")
AOO.Obs.Max.km2 <- setNames(AOO.Obs.Max.km2, "AOO.Obs.Max.km2")
#

### Determination of spatial pattern = "Small Patch"
Patch.Size.Avg.ha <- mean(TEM$AOO, na.rm = TRUE) / 1e4 # 5 h
Patch.Size.Med.ha <- median(TEM$AOO, na.rm = TRUE) / 1e4 # 2 h
Patch.Size.Max.ha <- max(TEM$AOO, na.rm = TRUE) / 1e4 # 181 ha
################################################  AOO Estimate and Inference ################################################
## Five TEM Projects (BAPID numbers: 4, 216, 4523, 4917, 6536,6605) recorded occupancy rates for this ecological community across five BGC units (ESSFdc2, IDFdk2, SBSdk, SBSmc2, SBSmk1). The rate of occupancy ranged from 0.03% to 1.02%.   THese rates of occupancy are multiplied by the min and max BGC range estimates to calculate an estimated area of occupancy below.
# BGC.Range.Mapped.Percent and BGC.Range.Mapped.km2 are defined above
AOO.Est.Min <- setNames((BGC.Range.Min.km2*0.0003),"AOO.Est.Min.km2")
AOO.Est.Max <- setNames((BGC.Range.Max.km2*0.0102),"AOO.Est.Max.km2")
#
#####################################################  NOO  #################################################
# Parameters
sep.dist.m <- 1000 #m
min.occ.size.km2 <- 0.0005 # 2 ha
### minimum NOO estimate while excluding TEM from the BWBS. See comments in Range extent. ###
valid.occurrences.min <- TEM %>% filter(BGC_ZONE != "BWBS" | is.na(BGC_ZONE))
# Filter polygons for min size
min_occ <- valid.occurrences.min[valid.occurrences.min$AOO.km2 >= min.occ.size.km2, ]
# remove empty records
min_occ_clean <- min_occ[!st_is_empty(min_occ), ]
# convert to terra object
min_occ_vect <- vect(min_occ_clean)
# create buffers
occ_buf <- buffer(min_occ_vect, width = (sep.dist.m/2))
# intersect buffers
adj_mat <- relate(occ_buf, occ_buf, relation = "intersects")
# Group buffers based on adjacency
n <- nrow(adj_mat)
group_id <- rep(NA_integer_, n)
current_group <- 0L

for (i in seq_len(n)) {
  if (is.na(group_id[i])) {
    current_group <- current_group + 1L
    to_visit <- i
    while (length(to_visit) > 0) {
      idx <- to_visit[1]
      to_visit <- to_visit[-1]
      if (is.na(group_id[idx])) {
        group_id[idx] <- current_group
        neighbors <- which(adj_mat[idx, ])
        if (length(neighbors)) {
          to_visit <- unique(c(to_visit, neighbors))
        }
      }
    }
  }
}
# Attach group IDs to buffers
occ_buf$GroupID <- group_id
# Dissolve buffers by GroupID
NOO.Groups.Min <- aggregate(occ_buf, by = "GroupID", dissolve = TRUE)
# convert back to sf
NOO.Groups.Min <- st_as_sf(NOO.Groups.Min)
### maximum NOO estimate including the BWBS. See comments in Range extent. ###
valid.occurrences.max <- TEM

# Filter polygons for min size
max_occ <- valid.occurrences.max[valid.occurrences.max$AOO.km2 >= min.occ.size.km2, ]
# remove empty records
max_occ_clean <- max_occ[!st_is_empty(max_occ), ]
# convert to terra object
max_occ_vect <- vect(max_occ_clean)
# create buffers
occ_buf <- buffer(max_occ_vect, width = (sep.dist.m/2))
# intersect buffers
adj_mat <- relate(occ_buf, occ_buf, relation = "intersects")
# Group buffers based on adjacency
n <- nrow(adj_mat)
group_id <- rep(NA_integer_, n)
current_group <- 0L

for (i in seq_len(n)) {
  if (is.na(group_id[i])) {
    current_group <- current_group + 1L
    to_visit <- i
    while (length(to_visit) > 0) {
      idx <- to_visit[1]
      to_visit <- to_visit[-1]
      if (is.na(group_id[idx])) {
        group_id[idx] <- current_group
        neighbors <- which(adj_mat[idx, ])
        if (length(neighbors)) {
          to_visit <- unique(c(to_visit, neighbors))
        }
      }
    }
  }
}

# Attach group IDs to buffers
occ_buf$GroupID <- group_id
# Dissolve buffers by GroupID
NOO.Groups.Max <- aggregate(occ_buf, by = "GroupID", dissolve = TRUE)
# convert back to sf
NOO.Groups.Max <- st_as_sf(NOO.Groups.Max)
#### Count NOO
NOO.Min.Count <- setNames(nrow(NOO.Groups.Min), "NOO.Min.Count")
NOO.Max.Count <- setNames(nrow(NOO.Groups.Max), "NOO.Max.Count")
#####################################################  TRENDS  #################################################
Human.Disturbance <- st_read(My.CEF.Human.Disturbance.path, "BC_CEF_HUMAN_DISTURBANCE_2023_fixed")

# Join (left join keeps all TEM rows).attaches all HD columns to TEM where they intersect.
Impacts <- st_join(
  TEM,
  Human.Disturbance,
  join = st_intersects,
  left = TRUE
)

st_write(Impacts, "TEM_joined_full.gpkg", layer = "TEM_join_full", delete_dsn = TRUE)
#################################################### WRITE SPATIAL #################################################################
# define output path
out.gpkg <- file.path(getwd(), "outputs", paste0(El.Sub.ID, "-csa-", format(Sys.time(), "%Y%m%d_%H%M"), ".gpkg"))

#write BEC Master points GPKG
st_write(BEC.Master.Focal.sf.buffer, out.gpkg, layer = "BECMaster")
#write TEM
st_write(TEM, out.gpkg, layer = "TEM")
# Write Range GPKG
st_write(BGC.MCP.Max, out.gpkg, layer = "BGC.MCP.Max")
st_write(BGC.Range.Min, out.gpkg, layer ="BGC.Range.Min")
st_write(BGC.Range.Max, out.gpkg, layer ="BGC.Range.max")
st_write(BGC.MCP.Min, out.gpkg, layer = "BGC.MCP.Min")
# Write NOO GPKG
st_write(NOO.Groups.Max, out.gpkg, layer = "NOO.Max")
st_write(NOO.Groups.Min, out.gpkg, layer = "NOO.Min")
# Write selected project boundaries and write ALL project boundaries (for ease of review)
st_write(TEI.proj.bound.all, out.gpkg, layer = "TEI.proj.bound.all")
st_write(TEI.proj.bound.min, out.gpkg, layer = "TEI.proj.bound.min")
st_write(TEI.proj.bound.max, out.gpkg, layer = "TEI.proj.bound.max")

############################################################### WRITE TABLES ####################################################
data.xlsx <- createWorkbook()

### Project Boundaries Data 
TEI.df <- st_drop_geometry(TEI.proj.bound.all)
addWorksheet(data.xlsx, "TEI Projects")
writeDataTable(data.xlsx, sheet = "TEI Projects", x = TEI.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")

### BEC Master Data 
addWorksheet(data.xlsx, "BEC.Master.Focal")
writeDataTable(data.xlsx, sheet = "BEC.Master.Focal", x = BEC.Master.Focal, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")

### TEM Data 
TEM.df <- st_drop_geometry(TEM)
addWorksheet(data.xlsx, "TEM")
writeDataTable(data.xlsx, sheet = "TEM", x = TEM.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")

### Geographic Summary
#BEC Master formatting
# add fields
BEC.Geographic.Summary <- BEC.Master.Focal %>%
  mutate(
    `Element Subnational ID` = El.Sub.ID,            # constant value in every row
    ID                      = PlotNumber,            # from BEC.Master.Focal$PlotNumber
    Source                  = "BEC_Master",          # constant string for all rows
    Notes                   = ""                     # blank notes
  ) %>%
# select and sort
  select(
    `Element Subnational ID`,
    ID,
    Latitude,
    Longitude,
    Source,
    Include,
    Notes
  )

#TEI formatting
TEI.Geographic.Summary <- TEM %>%
  # Select the attributes and rename/recode them to match Geographic.Summary
  mutate(
    `Element Subnational ID` = El.Sub.ID, # Match the constant from the original set
    ID                       = as.character(TEIS_ID),   # Use TEIS_ID for the new 'ID' column
    Source                   = 'TEI',     # Constant string for the new source
    Include                  = 'Yes',     # Constant 'Yes'
    Notes                    = "",        # Blank notes
  ) %>%
  # Select and sort
  select(
    `Element Subnational ID`,
    ID,
    Latitude,
    Longitude,
    Source,
    Include,
    Notes
  ) %>%
# drop geometry and Append
st_drop_geometry()
Geographic.Summary <- bind_rows(BEC.Geographic.Summary,TEI.Geographic.Summary)

# Write to Excel as a formatted table
addWorksheet(data.xlsx, "Geographic")

writeDataTable(
  wb         = data.xlsx,
  sheet      = "Geographic",
  x          = Geographic.Summary,
  startCol   = 1,
  startRow   = 1,
  tableStyle = "TableStyleLight8"
)


### Save workbook 
saveWorkbook(
  data.xlsx,
  file.path(getwd(), "outputs", paste0(El.Sub.ID, "-data-csa-", format(Sys.time(), "%Y%m%d_%H%M"), ".xlsx")),
  overwrite = TRUE
)
############################################################### WRITE SUMMARY ########################################################
summary.xlsx <- createWorkbook()

#write Range summary stats
addWorksheet(summary.xlsx, "Range")
writeData(summary.xlsx, "Range", as.data.frame(t(BGC.MCP.area.km2)), startRow = 1)
writeData(summary.xlsx, "Range", as.data.frame(t(BGC.Range.km2)), startRow = 3)
writeData(summary.xlsx, "Range", as.data.frame(t(BGC.Range.Mapped.km2)), startRow = 5)
writeData(summary.xlsx, "Range", as.data.frame(t()), startRow = 7)

#write AOO summary stats
addWorksheet(summary.xlsx, "AOO")
writeData(summary.xlsx, "AOO", as.data.frame(t(AOO.Obs.Min)),startRow = 1)
writeData(summary.xlsx, "AOO", as.data.frame(t(AOO.Obs.Max)),startRow = 3)
writeData(summary.xlsx, "AOO", as.data.frame(t(AOO.Est.Min)),startRow = 5)
writeData(summary.xlsx, "AOO", as.data.frame(t(AOO.Est.Max)),startRow = 7)
writeData(summary.xlsx, "AOO", as.data.frame(t(AOO.Inf.Min)),startRow = 9)
writeData(summary.xlsx, "AOO", as.data.frame(t(AOO.Inf.Max)),startRow = 11)

#write NOO summary stats
addWorksheet(summary.xlsx, "NOO")
writeData(summary.xlsx, "NOO", as.data.frame(t(NOO.Min.Count)),startRow = 1)
writeData(summary.xlsx, "NOO", as.data.frame(t(NOO.Max.Count)),startRow = 3)

#write CEF summary stats
addWorksheet(summary.xlsx, "CEF")
writeData(summary.xlsx, "CEF", CEF.Seral.Stage.Sums, startRow = 1)

#auto expand column widths
setColWidths(summary.xlsx, sheet = "Range", cols = 1:50, widths = "auto")
setColWidths(summary.xlsx, sheet = "AOO", cols = 1:50, widths = "auto")
setColWidths(summary.xlsx, sheet = "NOO", cols = 1:50, widths = "auto")
setColWidths(summary.xlsx, sheet = "CEF", cols = 1:50, widths = "auto")

#workbook is saved at the end of the next section

######################################################## FACTOR COMMENTS #######################################################
# assign text variables
Comments.Range.Extent <- glue("The total area of biogeoclimatic range of this ecological community is {round(BGC.Range.km2,2)} km2, and was calculated based on the sum of the area within the BWBSwk1, BWBSwk2, BWBSwk3 subzone variant polygons in BEC version 13 (citation).The estimated range extent is {round(BGC.MCP.area.km2,2)} km2, and was calculated based on the sum of the area of a minimum convex polygon plotted around the outer margins of the BWBSwk1, BWBSwk2, BWBSwk3 subzone variant polygons.")
Comments.AOO <- glue("The inferred Area of Occupancy for this ecological community is between {round(AOO.Inf.Min,2)} km2 and {round(AOO.Inf.Max,2)} km2. The AOO was first estimated by summing the proportion of area mapped as the focal ecological community in Terrestrial Ecosystem Information across the entire biogeoclimatic range of the focal ecological community. The minimum proportion of cover area was based on the sum of Terrestrial Ecosystem Information polygons mapped as in mature or old structural stages, and the maximum proportion was estimated based on the sum of these mature/old areas and additional polygons with no structural stage attribution (though still excluding polygons attributed to younger structural stages). The AOO was inferred by adjusting the estimates based on assumed rates of misattribution in Terrestrial Ecosystem Information Mapping. The minimum AOO inference was calculated by reducing the minimum estimate by 35% based on the assumption that up to 35% of the mapped area of focal site series are false positives. The maximum AOO inference was calculated by adjusting maximum AOO estimate based on the assumption that up to 35% of the area within the mapped biogeoclimatic range of this ecological community mapped as site series other than the focal sites series were incorrectly mapped and that up to 5% of this area is instead occupied by the focal ecological community")
Comments.NOO <- glue("There are between {NOO.Min.Count} and {NOO.Max.Count} occurrences of this ecological community. The minimum estimated number of occurrences were counted based on mapped focal sites series in Terrestrial Ecosystem Information polygons where structural stage is Mature or Old. The maximum estimated number of occurrences were counted based on mapped focal sites series in Terrestrial Ecosystem Information polygons where structural stage is Mature or Old and additional polygons where structural stage is Null (though still excluding polygons attributed to younger structural stages).  The separation distance between occurrences was {sep.dist.m} metres and the minimum occurrence size was {min.occ.size.km2} km2")
Comments.NOOGEI <- glue("")
Comments.AOOGEI <- glue("")
Comments.Env.Spe <- glue("")
Comments.Assigned.Threats <- glue("")
Comments.Calculated.Threats <- glue("")
Comments.Int.Vul <- glue("")
Comments.Short.Term.Trends <- glue("")
Comments.Long.Term.Trends <- glue("")
#Vector of variable names and row numbers
factor.comments <- c(
  "Comments.Range.Extent",
  "Comments.AOO",
  "Comments.NOO",
  "Comments.NOOGEI",
  "Comments.AOOGEI",
  "Comments.Env.Spe",
  "Comments.Assigned.Threats",
  "Comments.Calculated.Threats",
  "Comments.Int.Vul",
  "Comments.Short.Term.Trends",
  "Comments.Long.Term.Trends"
)
factor.comments.rows <- c(1,3,6,9,10,11,12,13,14,15,16)

addWorksheet(summary.xlsx, "Factor.Comments")

# Write variable names and values to specific rows to allow single paste into the rank calculator
for (i in seq_along(factor.comments)) {
  var_name <- factor.comments[i]
  var_value <- get(var_name, envir = .GlobalEnv)
  row_num <- factor.comments.rows[i]
  
  writeData(summary.xlsx, sheet = "Factor.Comments", x = var_name,  startCol = 1, startRow = row_num)
  writeData(summary.xlsx, sheet = "Factor.Comments", x = var_value, startCol = 2, startRow = row_num)
}
#auto expand column widths and apply text wrapping
wrapStyle <- createStyle(wrapText = TRUE)
addStyle(summary.xlsx, sheet = "Factor.Comments", style = wrapStyle, rows = 1:100, cols = 2:2, gridExpand = TRUE)
setColWidths(summary.xlsx, sheet = "Factor.Comments", cols = 1:50, widths = "auto")
# Save workbook
saveWorkbook(
  summary.xlsx,
  file.path(getwd(), "outputs", paste0(El.Sub.ID, "-data-csa-", format(Sys.time(), "%Y%m%d_%H%M"), ".xlsx")),
  overwrite = TRUE
)
#################################### UTILITIES #####################################################################
### Compare vetted BECMaster points with BGC units for Range Extent
# Keep only vetted plots
pts_in <- BEC.Master.Focal.sf %>%
  filter(Include == 1)
# Join overlapping features to select BGC units with valid plots in them
join_res <- st_join(
  pts_in, 
  BEC13 %>% select(MAP_LABEL), 
  join = st_intersects, 
  left = FALSE  # keep only points that intersect a polygon
)
st_write(join_res, out.gpkg, layer = "BECMasterTest")
# Extract MAP_LABELs (unique) and sort
map_labels <- join_res %>%
  pull(MAP_LABEL) %>%
  unique()

### ran this and confirmed that ALL BAPIDs with 'valid' mapcodes are in TEM 1 & 2
# For TEM1
bapid1 <- sf::st_read(dsn = TEM1.path, query = "SELECT BAPID FROM TEI_Long_Tbl", quiet = TRUE)
unique_bapid1 <- sort(unique(sf::st_drop_geometry(bapid1)$BAPID))

# For TEM2
bapid2 <- sf::st_read(dsn = TEM2.path, query = "SELECT BAPID FROM TEI_Long_Tbl", quiet = TRUE)
unique_bapid2 <- sort(unique(sf::st_drop_geometry(bapid2)$BAPID))

combined_bapid <- sort(unique(c(unique_bapid1, unique_bapid2)))

# valid bapids
check_vals <- c(1049, 1054, 4, 1048, 216, 135, 1046, 216, 233, 244)

# compare (all match)
result <- data.frame(
  value = check_vals,
  in_combined = ifelse(check_vals %in% combined_bapid, "yes", "no"),
  stringsAsFactors = FALSE
)

### Examine BAPIDS from LONG TABLE
# Unique values only (attribute-only query)
u <- st_read(TEI.Path, query = sprintf('SELECT DISTINCT "%s" FROM "%s"', "BAPID", "TEIS_Master_Long_Tbl"), quiet = TRUE)
overlapping.BAPIDS <- unique(TEI.proj.bound.all$BUSINESS_AREA_PROJECT_ID)

Missing.Projects <- setdiff(overlapping.BAPIDS, u$BAPID)

TEI.proj.bound.all$missing_bapid <- ifelse(
  TEI.proj.bound.all$BUSINESS_AREA_PROJECT_ID %in% Missing.Projects,
  "yes", "no"
)