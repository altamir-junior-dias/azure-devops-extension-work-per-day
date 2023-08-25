(() => {
    /* PUBLIC */
    window.WidgetConfiguration = {
        init: (WidgetHelpers) => {
            widgetHelpers = WidgetHelpers;
        },

        load: (widgetSettings, widgetConfigurationContext) => {
            context = widgetConfigurationContext;

            var settings = getSettings(widgetSettings);
            prepareControls(settings);

            return widgetHelpers.WidgetStatusHelper.Success();
        },

        save: (widgetSettings) => {
            return widgetHelpers.WidgetConfigurationSave.Valid(getSettingsToSave(true));
        }
    };

    /* PRIVATE */
    var context;
    var widgetHelpers;

    var $title = $('#title');
    var $source = $('#source');
    var $team = $('#team');
    var $iteration = $('#iteration');
    var $query = $('#query');
    var $startDate = $('#start-date');
    var $endDate = $('#end-date');
    var $memberField = $('#member-field');
    var $workField = $('#work-field');

    var $teamArea = $('#team-area');
    var $iterationArea = $('#iteration-area');
    var $queryArea = $('#query-area');
    var $startDateArea = $('#start-date-area');
    var $endDateArea = $('#end-date-area');

    var addQueryToSelect = (query, level) => {
        level = level ?? 0;

        if (query.isFolder ?? false) {
            $query.append($('<option>')
                .val(query.id)
                .html('&nbsp;&nbsp;'.repeat(level) + query.name)
                .attr('data-level', '0')
                .css('font-weight', 'bold')
                .attr('disabled', 'disabled'));

            if (query.children.length > 0)
            {
                query.children.forEach(innerQuery => {
                    addQueryToSelect(innerQuery, level + 1);
                });
            }

        } else {
            $query.append($('<option>')
                .val(query.id)
                .html('&nbsp;&nbsp;'.repeat(level) + query.name)
                .attr('data-level', level));
        }
    };

    var changeSettings = () => {
        settings = getSettingsToSave();

        var eventName = widgetHelpers.WidgetEvent.ConfigurationChange;
        var eventArgs = widgetHelpers.WidgetEvent.Args(settings);
        context.notify(eventName, eventArgs);
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

    var getSettingsToSave = () => {
        return {
            data: JSON.stringify({
                title: $title.val(),
                source: $source.val(),
                team: $team.val(),
                iteration: $iteration.val(),
                query: $query.val(),
                startDate: $startDate.val() != '' ? new Date($startDate.val()) : '',
                endDate: $endDate.val() != '' ? new Date($endDate.val()) : '',
                memberField: $memberField.val(),
                workField: $workField.val()
            })
        };
    };

    var prepareControls = (settings) => {
        var deferreds = [];

        deferreds.push(window.AzureDevOpsProxy.getTeams());
        deferreds.push(window.AzureDevOpsProxy.getSharedQueries());

        Promise.all(deferreds).then(result => {
            var teams = result[0];
            var queries = result[1];

            $query.append($('<option>'));

            queries.forEach(query => {
                addQueryToSelect(query);
            });

            $startDate.datepicker();
            $endDate.datepicker();

            $title.on('change', () => changeSettings());
            $source.on('change', () => {
                updateSources();
                changeSettings();
            });
            $team.on('change', () => updateIterations().then(_ => changeSettings()));
            $iteration.on('change', () => changeSettings());
            $query.on('change', () => changeSettings());
            $startDate.on('change', () => changeSettings());
            $endDate.on('change', () => changeSettings());
            $memberField.on('change', () => changeSettings());
            $workField.on('change', () => changeSettings());

            teams.forEach(team => $team.append($('<option>').val(team.id).html(team.name)));

            $title.val(settings.title);
            $source.val(settings.source);
            updateSources();
            $team.val(settings.team);
            updateIterations().then(_ => $iteration.val(settings.iteration));

            $query.val(settings.query);
            if (settings.startDate != '') {
                $startDate.datepicker('setDate', new Date(settings.startDate));
            }
            if (settings.endDate != '') {
                $endDate.datepicker('setDate', new Date(settings.endDate));
            }
            $memberField.val(settings.memberField);
            $workField.val(settings.workField);
        });
    };

    var updateIterations = () => {
        var deferred = $.Deferred();

        window.AzureDevOpsProxy.getIterations($team.val()).then(iterations => {
            $iteration.html('');

            $iteration.append(new Option('Current iteration', '0'));
            $iteration.append(new Option('Previous iteration', '-1'));

            iterations
                .sort((a, b) => a.startDate > b.startDate ? -1 : 1)
                .forEach(iteration => {
                    $iteration.append(new Option(iteration.name, iteration.id));
                });

            deferred.resolve();
        });

        return deferred.promise();
    };

    var updateSources = () => {
        if ($source.val() == 'iteration') {
            $teamArea.show();
            $iterationArea.show();
            $queryArea.hide();
            $startDateArea.hide();
            $endDateArea.hide();
        } else {
            $teamArea.hide();
            $iterationArea.hide();
            $queryArea.show();
            $startDateArea.show();
            $endDateArea.show();
        }
    };
})();