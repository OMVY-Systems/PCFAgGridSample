export const appConfig = {
    GET_URL: {

        // PARENT_DATA: "/api/data/v9.0/crfb2_projects?$select=crfb2_taskid,crfb2_parenttask,crfb2_taskname,crfb2_projectid,crfb2_percentagecomplete,crfb2_aplinestatus,crfb2_startdate,crfb2_enddate&$filter=crfb2_parenttask eq 'NA'",
        // FILTER_DATA: "/api/data/v9.0/crfb2_projects?$select=crfb2_taskid,crfb2_parenttask,crfb2_taskname,crfb2_projectid,crfb2_percentagecomplete,crfb2_aplinestatus,crfb2_startdate,crfb2_enddate&$filter=crfb2_parenttask eq '",
        // APLINE_LOOKUP: "/api/data/v9.0/EntityDefinitions(LogicalName='crfb2_project')/Attributes(LogicalName='crfb2_aplinestatus')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"

        PARENT_DATA: "/api/data/v9.0/crfb2_projects?$select=crfb2_taskid,crfb2_taskname,crfb2_projectid,crfb2_percentagecomplete,crfb2_aplinestatus,crfb2_startdate,crfb2_enddate,_crfb2_parent_value&$filter=_crfb2_parent_value eq null",
        FILTER_DATA: "/api/data/v9.0/crfb2_projects?$select=crfb2_taskid,crfb2_taskname,crfb2_projectid,crfb2_percentagecomplete,crfb2_aplinestatus,crfb2_startdate,crfb2_enddate,_crfb2_parent_value&$filter=_crfb2_parent_value eq ",
        APLINE_LOOKUP: "/api/data/v9.0/EntityDefinitions(LogicalName='crfb2_project')/Attributes(LogicalName='crfb2_aplinestatus')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"
    },

    SCHEMA: {
        ENTITY_NAME_FOR_UPDATE: "crfb2_project",
        ENTITY_NAME_FOR_BATCH_UPDATE: "crfb2_projects"
    }

}