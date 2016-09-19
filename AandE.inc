//This section of the configuration file sets where the user guide is located
SET vUserGuideLocation = 'User guides\QlikView User Guide - A&E App.pdf';
SET vInformationPackLocation = 'Information Packs\Information Pack - A&E Application.pdf';

//This section of the configuration file sets where the required files are located
//This can either contain a relative or an absolute path

SET vTablenameAandECurrentWaits = 'CDI\AandE.AandECurrentWaits.xlsx';
SET vTablenameAccidentAndEmergency = 'CDI\AandE.AccidentAndEmergency.csv';
SET vTablenameAttendanceDisposalCode = 'CDI\AandE.ref_AttendanceDisposalCode.xlsx';
SET vTablenameAttendanceCategoryCode = 'CDI\AandE.ref_AttendanceCategoryCode.xlsx';
SET vTablenamePatientTypeCode = 'CDI\AandE.ref_PatientTypeCode.xlsx';
SET vTablenameSourceOfReferralAECode = 'CDI\AandE.ref_SourceOfReferralAECode.xlsx';
SET vTablenameBreachReasonCode = 'CDI\AandE.ref_BreachReasonCode.xlsx';
SET vTablenameArrivalModeCode = 'CDI\AandE.ref_ArrivalModeCode.xlsx';
SET vTablenameConsultantCode = 'CDI\AandE.ref_ConsultantCode.xlsx';
SET vTablenameSitReps = 'CDI\Peers.csv';
SET vTablenamePeers = 'CDI\Peers.csv';

//This section of the configuration file sets what format the required files have been saved as
//This can either contain a relative or an absolute path

SET vFiletypeAandECurrentWaits = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeAccidentAndEmergency = 'csv'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeAttendanceDisposalCode = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeAttendanceCategoryCode = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypePatientTypeCode = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeSourceOfReferralAECode = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeBreachReasonCode = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeArrivalModeCode = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeConsultantCode = 'xlsx'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypeSitReps = 'csv'; //Enter either 'xlsx', 'xls' or 'csv' here
SET vFiletypePeers = 'csv'; //Enter either 'xlsx', 'xls' or 'csv' here

//This section of the configuration file will replace the logo on each application with the 
//corresponding user

SET vApplicationLogo = 'CDI\UH-logo-black-and-blue.jpg';

//This Section of the configuration file will add the recipient and the subject for the Email Contact us button

SET vEmailRecipient = 'Support@Draperanddash.com';
SET vEmailSubject = 'Application Support';

//This Section of the configuration file will help the AutoMail Stuff work...
SET vMailPath = =QvWorkPath;
SET vActiveSheetName = =((REPLACE(REPLACE($(=Subfield(GetActiveSheetId(),'\',2)),':',''),' ','') ) );
SET vSendMail = 0

