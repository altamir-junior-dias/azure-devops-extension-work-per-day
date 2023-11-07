(() => {
    /* PUBLIC */
    window.AzureDevOpsProxy = {
        getBacklogOld: (teamId, backlogId, backlogName) => {
            var deferred = $.Deferred();

            var context = {
                projectId: VSS.getWebContext().project.id,
                teamId: teamId
            };

            var deferreds = [];

            deferreds.push(tfsWebAPIClient.getBoard(context, backlogName));
            deferreds.push(tfsWebAPIClient.getBacklog(context, backlogId));

            Promise.all(deferreds).then(result => {
                var board = result[0];
                var backlog = result[1];

                deferreds = [];
                backlog.workItemTypes.forEach(workItemType => deferreds.push(window.AzureDevOpsProxy.getWorkItemType(workItemType.name)));

                Promise.all(deferreds).then(workItemTypes => {
                    var result = {
                        lanes: board.rows.map(row => {
                            return {
                                id: row.id,
                                name: row.name
                            };
                        }),
                        columns: board.columns
                            .map(column => {
                                return {
                                    id: column.id,
                                    name: column.name
                                };
                            }),
                        workItemTypes: backlog.workItemTypes.map(workItemType => {
                            return {
                                name: workItemType.name,
                                color: workItemTypes.find(wit => wit.name == workItemType.name).color
                            };
                        })
                    };

                    var states = board.columns
                        .map(column => Object.values(column.stateMappings)
                        .filter((value, index, array) => array.indexOf(value) === index ))
                        .join().split(',').filter((value, index, array) => array.indexOf(value) === index);
                    var workItemTypes = backlog.workItemTypes.map(workItemType => workItemType.name);

                    var query = {
                        wiql: 'SELECT [System.Id], [System.Title], [System.BoardColumn], [System.State], [System.WorkItemType], [System.CreatedDate], [Microsoft.VSTS.Common.ClosedDate] ' + 
                            'FROM WorkItems ' + 
                            'WHERE [System.State] in (' + states.map(state => "'" + state + "'").join(',') + ') ' +
                            '  AND [System.WorkItemType] in (' + workItemTypes.map(workItemType => "'" + workItemType + "'").join(',') + ')' +
                            '  AND [System.BoardColumn] <> \'\'',
                        type: '1'
                    };

                    window.AzureDevOpsProxy.getItemsFromQuery(query).then(items => {
                        //var inProgressColumns = board.columns
                            //.filter(column => column.columnType == 1)
                            //.map(column => column.name);

                        var ids = items
                            //.filter(item => inProgressColumns.findIndex(column => column == item['System.BoardColumn']) > -1)
                            .map(item => item.id);

                        getRevisionsFromItemsv2(ids).then(revisions => {
                            result.items = items.map(item => {
                                item.revisions = revisions.filter(revision => revision.id == item.id);

                                return item;
                            });
    
                            deferred.resolve(result);
                        });
                    });
                });
            });

            return deferred.promise();
        },

        getBacklog: (teamId, backlogId, backlogName) => {
            var deferred = $.Deferred();

            var context = {
                projectId: VSS.getWebContext().project.id,
                teamId: teamId
            };

            var deferreds = [];

            deferreds.push(tfsWebAPIClient.getBoard(context, backlogName));
            deferreds.push(tfsWebAPIClient.getBacklog(context, backlogId));

            Promise.all(deferreds).then(result => {
                var board = result[0];
                var backlog = result[1];

                deferreds = [];
                backlog.workItemTypes.forEach(workItemType => deferreds.push(window.AzureDevOpsProxy.getWorkItemType(workItemType.name)));

                Promise.all(deferreds).then(workItemTypes => {
                    var result = {
                        lanes: board.rows.map(row => {
                            return {
                                id: row.id,
                                name: row.name
                            };
                        }),
                        columns: board.columns
                            .map(column => {
                                return {
                                    id: column.id,
                                    name: column.name
                                };
                            }),
                        workItemTypes: backlog.workItemTypes.map(workItemType => {
                            return {
                                name: workItemType.name,
                                color: workItemTypes.find(wit => wit.name == workItemType.name).color
                            };
                        })
                    };

                    deferred.resolve(result);
                });
            });

            return deferred.promise();
        },

        getBacklogs: (teamId) => {
            var deferred = $.Deferred();

            var context = {
                projectId: VSS.getWebContext().project.id,
                teamId: teamId
            };

            tfsWebAPIClient.getBacklogs(context).then(backlogs => {
                backlogs.sort((a, b) => a.rank > b.rank ? -1 : a.rank < b.rank ? 1 : 0)
                deferred.resolve(backlogs
                    .filter(backlog => backlog.type != 2)
                    .map(backlog => {
                        return {
                            id: backlog.id,
                            name: backlog.name
                        };
                    }));
            });

            return deferred.promise();
        },

        getChartService: () => {
            var deferred = $.Deferred();

            chartsServices.ChartsService.getService().then((chartService) => {
                deferred.resolve(chartService);
            });

            return deferred.promise();
        },

        getCurrentIterarion: (teamId, skip) => {
            var deferred = $.Deferred();
    
            var webContext = VSS.getWebContext();
            var teamContext = { projectId: webContext.project.id, teamId: teamId || webContext.team.id, project: "", team: "" };
    
            tfsWebAPIClient.getTeamIterations(teamContext, "current").then((items) => {
                var iterations = items.map(i => { 
                    return { 
                        id: i.id, 
                        name: i.name, 
                        path: i.path, 
                        startDate: i.attributes.startDate, 
                        endDate: i.attributes.finishDate 
                    }; 
                });
    
                if (skip === undefined) {
                    deferred.resolve(iterations[0]);
                } else {
                    var currentIteration = iterations[0];
    
                    window.AzureDevOpsProxy.getIterations(teamId).then(iterations => {
                        var orderedIterations = iterations.sort((a, b) => { return a.startDate > a.endDate ? 1 : 1; });
    
                        var index = orderedIterations.findIndex(interation => interation.id == currentIteration.id) + skip;
    
                        if (index < 0) {
                            deferred.resolve(iterations[0]);
                        } else if (index > iterations.length - 1) {
                            deferred.resolve(iterations[iterations.length -1]);
                        } else {
                            deferred.resolve(iterations[index]);
                        }
                    });
                }
            });
    
            return deferred.promise();
        },

        getCurrentUser: () => { 
            return VSS.getWebContext().user;
        },

        getExtensionData: (data, defaultValue) => {
            var deferred = $.Deferred();

            extensionDataService.getValue(data)
                .then((value) => {
                    deferred.resolve(value ?? defaultValue);
                },
                () => {
                    deferred.resolve(defaultValue);
                });

            return deferred.promise();
        },

        getItems: (query, asOf) => {
            var deferred = $.Deferred();
    
            var projectId = VSS.getWebContext().project.id;
            witClient.queryByWiql({ query: getCleanedQuery(query, asOf) }, projectId).then((result) => {
                let ids = result.workItems.map(r => r.id);
    
                if (ids.length > 0) {
                    getWorkItemsById(ids, getQueryFields(query), asOf).then((workItems) => deferred.resolve(workItems));
                } else {
                    deferred.resolve([]);
                }
            });
    
            return deferred.promise();
        },

        getItemsFromQuery: (query, withRevisions) => {
            withRevisions = withRevisions ?? false;

            var deferred = $.Deferred();
    
            var projectId = VSS.getWebContext().project.id;
            witClient.queryByWiql({ query: getCleanedQuery(query.wiql) }, projectId).then((result) => {
                var ids = [];

                if (query.type == "1") {
                    ids = result.workItems.map(r => r.id);

                } else {
                    var sources = result.workItemRelations.filter(r => r.source != null).map(r => r.source.id)
                    var targets = result.workItemRelations.map(r => r.target.id);
        
                    var uniqueSources = sources.filter(function (value, index, self) { return self.indexOf(value) === index; });
                    var uniqueTargets = targets.filter(function (value, index, self) { return self.indexOf(value) === index; });
        
                    ids = uniqueSources.concat(uniqueTargets);                
                }
    
                if (ids.length > 0) {
                    var deferreds = [];

                    deferreds.push(getWorkItemsById(ids, getQueryFields(query.wiql)));

                    if (withRevisions) {
                        deferreds.push(getRevisionsFromItems(ids));
                    }
                    
                    Promise.all(deferreds).then(results => {
                        var workItems = results[0];
                        var revisions = results[1];

                        if (query.type == "1") {
                            if (withRevisions) {
                                workItems = workItems.map(workItem => {
                                    workItem.revisions = revisions.filter(revision => revision.id == workItem.id);

                                    return workItem;
                                });
                            }

                            deferred.resolve(workItems);

                        } else {
                            var items = result.workItemRelations
                                .filter(r => r.source == null)
                                .map(r => {
                                    var parent = workItems.find(wi => wi.id == r.target.id);
        
                                    if (withRevisions) {
                                        parent.revisions = revisions.filter(revision => revision.id == parent.id);
                                    }

                                    parent.children = result.workItemRelations
                                        .filter(s => s.source != null && s.source.id == r.target.id)
                                        .map(s => {
                                            var child = workItems.find(wi => wi.id == s.target.id);

                                            if (withRevisions) {
                                                child.revisions = revisions.filter(revision => revision.id == child.id);
                                            }

                                            return child;
                                        });
        
                                    return parent;
                                }); 

                            deferred.resolve(items);
                        }
                    });
                } else {
                    deferred.resolve([]);
                }
            });
    
            return deferred.promise();
        },

        getItemRevisions: (id, skip, revisions) => {
            var deferred = $.Deferred();

            skip = skip || 0;
            revisions = revisions || [];

            witClient.getRevisions(id, 200, skip).then(nextRevisions => {
                if (nextRevisions.length == 200) {
                    skip += 200;

                    window.AzureDevOpsProxy.getItemRevisions(id, skip, revisions.concat(nextRevisions)).then(nextRevisions => {
                        deferred.resolve(nextRevisions);
                    });
                } else {
                    deferred.resolve(revisions.concat(nextRevisions))
                }
            });

            return deferred.promise();
        },
        
        getIteration: (iteration, team) => {
            var deferred = $.Deferred();
    
            var webContext = VSS.getWebContext();
            var teamContext = { projectId: webContext.project.id, teamId: team ?? webContext.team.id, project: "", team: "" };
    
            tfsWebAPIClient.getTeamIteration(teamContext, iteration).then(data => {
                deferred.resolve({ 
                    id: data.id, 
                    name: data.name, 
                    path: data.path, 
                    startDate: data.attributes.startDate, 
                    endDate: data.attributes.finishDate 
                });    
            });
    
            return deferred.promise();
        },

        getIterations: (teamId) => {
            var deferred = $.Deferred();
            var webContext = VSS.getWebContext();

            var teamContext = { projectId: webContext.project.id, teamId: teamId ?? webContext.team.id, project: "", team: "" };
    
            tfsWebAPIClient.getTeamIterations(teamContext).then((items) => {
                var iterations = items.map(i => { 
                    return { 
                        id: i.id, 
                        name: i.name, 
                        path: i.path, 
                        startDate: i.attributes.startDate, 
                        endDate: i.attributes.finishDate 
                    }; 
                });
    
                deferred.resolve(iterations);
            });
    
            return deferred.promise();
        },

        getIterationItems: (teamId, iterationId) => {
            var deferred = $.Deferred();

            window.AzureDevOpsProxy.getTeamAreas(teamId).then(areas => {
                var areasFilter = areas.map(area => '[System.AreaPath]' + (area.includeChildren ? ' under ' : ' = ') + '\'' + area.value + '\'').join(' OR ');

                window.AzureDevOpsProxy.getIteration(iterationId, teamId).then(iteration => {
                    var query = 
                        'SELECT [System.Id], [System.Title] ' + 
                        'FROM WorkItems ' + 
                        'WHERE [System.IterationPath] = \'' + iteration.path + '\' ' +
                        '  AND (' + areasFilter + ') ';

                    var startDate = new Date(iteration.startDate);
                    var endDate = new Date(iteration.endDate) > new Date() ? new Date() : new Date(iteration.endDate);

                    var deferreds = [];
                    var date = startDate

                    for (var date = startDate; date <= endDate; date = new Date(date.setDate(date.getDate() + 1))) {
                        deferreds.push(witClient.queryByWiql({ query: getCleanedQuery(query, date) }, VSS.getWebContext().project.id));
                    }

                    Promise.all(deferreds).then(result => {
                        var ids = result
                            .map(r => r.workItems.map(workItem => workItem.id))
                            .flat(1)
                            .sort()
                            .filter((value, index, self) => self.indexOf(value) === index);
                        
                        query = 
                            'SELECT [System.Id], [System.Title] ' + 
                            'FROM WorkItems ' + 
                            'WHERE [System.Id] IN (' + ids.join(',') + ')';

                        window.AzureDevOpsProxy.getItems(query).then(result => {
                            ids = result.map(r => r.id);

                            getRevisionsFromItems(ids).then(revisions => {
                                deferred.resolve(ids.map(id => {
                                    return {
                                        id: id,
                                        revisions: revisions.filter(revision => revision.id == id)
                                    }
                                }));
                            });
                        });
                    });
                });
            });

            return deferred.promise();
        },
    
        getParentChildren: (query, getChildRevisions) => {
            var deferred = $.Deferred();
    
            var projectId = VSS.getWebContext().project.id;
    
            witClient.queryByWiql({ query: getCleanedQuery(query) }, projectId).then((queryResult) => {
                var sources = queryResult.workItemRelations.filter(r => r.source != null).map(r => r.source.id)
                var targets = queryResult.workItemRelations.map(r => r.target.id);
    
                var uniqueSources = sources.filter(function (value, index, self) { return self.indexOf(value) === index; });
                var uniqueTargets = targets.filter(function (value, index, self) { return self.indexOf(value) === index; });
    
                var ids = uniqueSources.concat(uniqueTargets);
    
                var deferreds = [];
    
                deferreds.push(getWorkItemsById(ids, getQueryFields(query)));

                if (getChildRevisions) {
                    deferreds.push(getRevisionsFromItems(ids));
                }
    
                Promise.all(deferreds).then(result => {
                    var workItems = result[0];
    
                    var parents = queryResult.workItemRelations
                        .filter(r => r.source == null)
                        .map(r => {
                            var parent = workItems.find(wi => wi.id == r.target.id).fields;
    
                            parent.id = r.target.id;
                            parent.children = queryResult.workItemRelations
                                .filter(s => s.source != null && s.source.id == r.target.id)
                                .map(s => {
                                    var child = workItems.find(wi => wi.id == s.target.id).fields;
                                    child.id = s.target.id;
    
                                    if (getChildRevisions) {
                                        child.revisions = result[1].filter(r => r.id == child.id);
                                    }
    
                                    return child; 
                                });
    
                            return parent;
                        });
        
                    deferred.resolve(parents);
                });
            });
    
            return deferred.promise();
        },

        getQuery: (queryPath) => {
            var deferred = $.Deferred();
            var webContext = VSS.getWebContext();
    
            var sendResult = (deferred, result) => {
                result.children.sort((a, b) => {
                    var result = a.isFolder && !b.isFolder ? -1 : !a.isFolder && b.isFolder ? 1 : 0;
    
                    if (result == 0) {
                        result = a.path > b.path ? 1 : a.path < b.path ? -1 : 0;
                    }
    
                    return result;
                });
    
                deferred.resolve(result);
            };
    
            witRestClient.getQuery(webContext.project.id, queryPath, 'all', 1).then(function (query) {
                var children = query.children ?? [];
                var childrenQueries = children.filter(query => !(query.hasChildren ?? false));
                var childrenFolders = children.filter(query => query.hasChildren ?? false);

                var result = {
                    id: query.id,
                    name: query.name,
                    path: query.path,
                    isFolder: query.isFolder ?? false,
                    children: childrenQueries.map(childQuery => {
                        return {
                            id: childQuery.id,
                            name: childQuery.name,
                            path: childQuery.path,
                            isFolder: childQuery.isFolder ?? false,
                            children: []
                        };
                    })
                };

                if (childrenFolders.length > 0) {
                    var deferreds = [];

                    childrenFolders.forEach(child => deferreds.push(window.AzureDevOpsProxy.getQuery(child.id)));

                    Promise.all(deferreds).then(queries => {
                        queries.forEach(childQuery => {
                            result.children.push({
                                id: query.id,
                                name: childQuery.name,
                                path: childQuery.path,
                                isFolder: childQuery.isFolder ?? false,
                                children: childQuery.children
                            });
                        });

                        sendResult(deferred, result);
                    });

                } else {

                    sendResult(deferred, result);
                }
            });
    
            return deferred.promise();
        },

        getQueryWiql: (queryId) => {
            var deferred = $.Deferred();
            var webContext = VSS.getWebContext();

            witRestClient.getQuery(webContext.project.id, queryId, 'all', 1).then(function (query) {
                deferred.resolve({
                    wiql: query.wiql,
                    type: query.queryType,
                    name: query.name
                });
            });

            return deferred.promise();
        },

        getSharedQueries: () => {
            var deferred = $.Deferred();

            window.AzureDevOpsProxy.getQuery('Shared Queries').then(function (query) {
                deferred.resolve(query.children);
            });

            return deferred.promise();
        },

        getTeams: () => {
            var deferred = $.Deferred();

            var projectId = VSS.getWebContext().project.id;

            tfsCoreRestClient.getTeams(projectId).then(teams => {
                deferred.resolve(teams.map(team => {
                    return {
                        id: team.id,
                        name: team.name
                    };
                }));
            });

            return deferred.promise();
        },

        getTeam: (teamId) => {
            var deferred = $.Deferred();

            var projectId = VSS.getWebContext().project.id;

            tfsCoreRestClient.getTeam(projectId, teamId).then(team => {
                deferred.resolve(team);
            });

            return deferred.promise();
        },

        getTeamAreas: (teamId) => {
            var deferred = $.Deferred();

            var webContext = VSS.getWebContext();
            var teamContext = { projectId: webContext.project.id, teamId: teamId || webContext.team.id, project: "", team: "" };

            tfsWebAPIClient.getTeamFieldValues(teamContext).then((fields) => {
                deferred.resolve(fields.values);
            });

            return deferred.promise();
        },

        getWorkItemType : (referenceName) => {
            var deferred = $.Deferred();
    
            var webContext = VSS.getWebContext();
            witClient.getWorkItemType(webContext.project.id, referenceName).then(item => {
                deferred.resolve({
                    referenceName: item.referenceName,
                    name: item.name,
                    color: item.color
                });
            });
    
            return deferred.promise();
        },

        getWorkItemTypeFields : (referenceName) => {
            var deferred = $.Deferred();
    
            var webContext = VSS.getWebContext();
            witClient.getWorkItemType(webContext.project.id, referenceName).then(wotkItemType => {
                var fields = wotkItemType.fields.map(field => {
                    return {
                        referenceName: field.referenceName,
                        name: field.name,
                        type: getFieldType(field.referenceName)
                    };
                });

                deferred.resolve(fields);
            });
    
            return deferred.promise();
        },

        getWorkItemTypes: () => {
            var deferred = $.Deferred();
    
            var webContext = VSS.getWebContext();
            witClient.getWorkItemTypes(webContext.project.id).then(items => {
                var workItemType = items.map(item => {
                    return {
                        referenceName: item.referenceName,
                        name: item.name
                    };
                });
    
                deferred.resolve(workItemType);
            });
    
            return deferred.promise();
        },

        init: (VSSService, TFSWorkItemTrackingRestClient, TFSWorkRestClient, TFSCoreRestClient, ExtensionDataService, ChartsServices) => {
            witClient = VSSService?.getCollectionClient(TFSWorkItemTrackingRestClient.WorkItemTrackingHttpClient);
            witRestClient = TFSWorkItemTrackingRestClient.getClient();
            tfsWebAPIClient = TFSWorkRestClient?.getClient();
            tfsCoreRestClient = TFSCoreRestClient.getClient();
            extensionDataService = ExtensionDataService;
            chartsServices = ChartsServices;

            var webContext = VSS.getWebContext();
            var deferred = witClient?.getFields !== undefined ? witClient?.getFields(webContext.project.id) : witClient?.getWorkItemFields(webContext.project.id);
            deferred.then(fields => {
                projectFields = fields.map(field => {
                    return {
                        referenceName: field.referenceName,
                        type: field.type
                    };
                });
            });
        },

        saveExtensionData: (data, value) => {
            var deferred = $.Deferred();

            extensionDataService.setValue(data, value).then(() => {
                deferred.resolve();
            });

            return deferred.promise();
        }
    };

    /* PRIVATE */
    var witClient;
    var witRestClient;
    var tfsWebAPIClient;
    var tfsCoreRestClient;
    var extensionDataService;
    var chartsServices;

    var projectFields;

    var getCleanedQuery = (query, asOf) => {
        let queryFields = query
            .substring(query.toUpperCase().indexOf("SELECT") + 7, query.toUpperCase().indexOf("FROM"))
            .trim();

        let cleanedQuery = query
            .replace(queryFields, '[System.Id]');

        if (asOf !== undefined && asOf != null)
        {
            cleanedQuery += 'ASOF \'' + asOf.toISOString().split('T')[0] + '\'';
        }

        return cleanedQuery;
    };

    var getQueryFields = (query) => {
        var queryFields = query
            .substring(query.toUpperCase().indexOf("SELECT") + 7, query.toUpperCase().indexOf("FROM"))
            .trim();

        var fields = queryFields
            .toUpperCase()
            .replace('[SYSTEM.ID]', '')
            .trim()
            .split(',')
            .filter(f => f.trim() != '')
            .map(f => f.replace('[', '').replace(']', '').trim());

        return fields;
    };

    var getRevisionsFromItem = (id) => {
        let deferred = $.Deferred();

        witClient.getRevisions(id).then(revisions => {
            deferred.resolve(revisions);
        });

        return deferred.promise();
    };

    var getRevisionsFromItems = (ids) => {
        var deferred = $.Deferred();

        var deferreds = [];

        ids.forEach(id => deferreds.push(getRevisionsFromItem(id)));

        Promise.all(deferreds).then(result => {
            deferred.resolve([].concat.apply([], result));
        });

        return deferred.promise();
    };

    var getRevisionsFromItemsv2 = (ids) => {
        var deferred = $.Deferred();

        if (ids.length > 100) {
            let pack = ids.slice(0, 100);

            var deferreds = [];
            pack.forEach(id => deferreds.push(getRevisionsFromItem(id)));

            Promise.all(deferreds).then(result => {
                var revisions = [].concat.apply([], result);
                var newPack = ids.slice(101);

                getRevisionsFromItemsv2(newPack).then(result => {
                    deferred(revisions.concat.apply([], result));
                });
            });
    
        } else {
            var deferreds = [];
            ids.forEach(id => deferreds.push(getRevisionsFromItem(id)));

            Promise.all(deferreds).then(result => {
                var revisions = [].concat.apply([], result);

                deferred.resolve(revisions);
            });
        }

        return deferred.promise();
    }

    var getWorkItemsById = (ids, fields, asOf) => {
        fields = fields || [];

        var deferred = $.Deferred();

        var deferreds = [];

        for(let i = 0; i < ids.length; i += 50)
        {
            let pack = ids.slice(i, i + 50);

            deferreds.push(witClient.getWorkItems(pack, fields, asOf));
        }

        Promise.all(deferreds).then(result => {
            var items = [].concat.apply([], result).map(item => {
                var itemWithFields = item.fields;
                itemWithFields.id = item.id;

                return itemWithFields;
            });

            deferred.resolve(items);
        });

        return deferred.promise();
    };

    var getFieldType = (referenceName) => {
        return projectFields.find(field => field.referenceName == referenceName)?.type;
    };
})();
