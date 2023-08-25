(() => {
    /* PUBLIC */
    window.Widget = {
        load: (widgetSettings) => {
            var settings = getSettings(widgetSettings);

            $title.text(settings.title);

            getData(settings).then(data => prepareWidget(data));

            return window.WidgetHelpers.WidgetStatusHelper.Success();
        }
    };

    /* PRIVATE */
    var $title = $('#title');
    var $subTitle = $('#sub-title');
    
    var $titleArea = $('#title-area');
    var $widgetBody = $('#widget-body');

    var getData = (settings) => {
        var deferred = $.Deferred();

        if (settings.source == 'iteration') {
            getIteration(settings).then(iteration => {
                $subTitle.text(iteration.name);

                var teamId = settings.team;
                window.AzureDevOpsProxy.getIterationItems(teamId, iteration.id).then(items => {
                    deferred.resolve({
                        items: items,
                        startDate: iteration.startDate,
                        endDate: iteration.endDate,
                        memberField: settings.memberField,
                        workField: settings.workField
                    });
                });
            });
        } else {
            window.AzureDevOpsProxy.getQueryWiql(settings.query).then(query => {
                $subTitle.text(query.name);

                window.AzureDevOpsProxy.getItemsFromQuery(query, true).then(items => {
                    deferred.resolve({
                        items: items
                            .filter(item => item.children !== undefined)
                            .map(item => item.children)
                            .flat(1)
                            .concat(items),
                        startDate: new Date(settings.startDate),
                        endDate: new Date(settings.endDate),
                        memberField: settings.memberField,
                        workField: settings.workField
                    });
                });
            });
        }

        return deferred.promise();
    };

    var getIteration = (settings) => {
        var deferred = $.Deferred();

        if (settings.iteration == "0") {
            window.AzureDevOpsProxy.getCurrentIterarion(settings.team).then(iteration => {
                deferred.resolve(iteration);
            });
        } else if (settings.iteration == "-1") {
            window.AzureDevOpsProxy.getCurrentIterarion(settings.team, -1).then(iteration => {
                deferred.resolve(iteration);
            });
        } else {
            window.AzureDevOpsProxy.getIteration(settings.iteration).then(iteration => {
                deferred.resolve(iteration);
            });
        }

        return deferred.promise();
    };

    var getSettings = (widgetSettings) => {
        var settings = JSON.parse(widgetSettings.customSettings.data);

        return {
            title: settings?.title ?? 'Work Per Day',
            source: settings?.source ?? 'iteration',
            team: settings?.team ?? VSS.getWebContext().team.id,
            iteration: settings?.iteration ?? 0,
            query: settings?.query ?? '',
            startDate: settings?.startDate ?? '',
            endDate: settings?.endDate ?? '',
            memberField: settings?.memberField ?? 'System.ChangedBy',
            workField: settings?.workField ?? 'Microsoft.VSTS.Scheduling.CompletedWork'
        };
    };

    var prepareChart = (chart) => {
        window.AzureDevOpsProxy.getChartService().then((chartService) => {
            var $container = $('#chart');
            var chartOptions = { 
                'hostOptions': { 
                    'height': chart.height, 
                    'width': chart.width
                },
                'chartType': 'table',
                "xAxis": {
                    "labelValues": chart.xAxis
                },
                "yAxis": {
                    "labelValues": chart.yAxis
                },
                "series": chart.series
            }

            chartService.createChart($container, chartOptions);
        });
    };

    var prepareSeries = (dates, members, changes) => {
        var series = [];
        var column = 0;

        dates.forEach(date => {
            var line = 0;

            series.push({
                name: date,
                data: []
            });

            members.forEach(member => {
                var indexMember = changes.findIndex(v => v.member.uniqueName == member.uniqueName);
                var indexDate = changes[indexMember].dates.findIndex(v => v.date == date);
                var total = changes[indexMember].dates[indexDate].total;

                series[series.length -1].data.push([column, line, total])

                line = line + 1;
            });

            column = column + 1;
        });

        return series;
    };

    var prepareWidget = (data) => {
        var members = [];
        var dates = [];
        var changes = [];

        data.items.forEach(item => {
            var currentWork = 0;
    
            item.revisions.forEach(revision => {
                var changedDate = new Date(revision.fields['System.ChangedDate']);
                var member = revision.fields[data.memberField];
                var work = revision.fields[data.workField] || 0;

                var workHasChange = currentWork != work;
                var changeWithinTheSprint = changedDate >= data.startDate && changedDate <= data.endDate;

                if (workHasChange && changeWithinTheSprint && member != null)
                {
                    var workChanged = work - currentWork;
                    var onlyDate = changedDate.toISOString().split('T')[0];
                    var newDate = dates.findIndex(d => d == onlyDate) == -1;
                    var newMember = members.findIndex(m => m.uniqueName == member.uniqueName) == -1;
                    var index = changes.findIndex(c => c.member == member.uniqueName && c.date == onlyDate);

                    if (newDate) {
                        dates.push(onlyDate);
                    }

                    if (newMember) {
                        members.push(member);
                    }

                    if (index == -1) {
                        changes.push({
                            member: member.uniqueName,
                            date: onlyDate,
                            total: 0
                        });

                        index = changes.length - 1;
                    }

                    changes[index].total = changes[index].total + workChanged;

                    currentWork = work;
                }
            });
        });

        dates.sort();
        members.sort((a, b) => a.displayName > b.displayName ? 1 : -1);

        var result = [];

        members.forEach(member => {
            var line = {
                member: member,
                dates: []
            };

            dates.forEach(date => {
                line.dates.push({
                    date: date,
                    total: changes
                        .filter(c => c.member == member.uniqueName && c.date == date)
                        .map(c => c.total)
                        .reduce((a, b) => a + b, 0)
                });
            });

            result.push(line);
        });

        prepareChart({
            height: $widgetBody.height() - $titleArea.height(),
            width: $widgetBody.width(),
            xAxis: dates.map(d => new Date(d).toLocaleDateString()),
            yAxis: members.map(m => m.displayName),
            series: prepareSeries(dates, members, result)
        });
    }
})();