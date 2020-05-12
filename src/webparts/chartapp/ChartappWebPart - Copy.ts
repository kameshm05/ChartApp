import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ChartappWebPart.module.scss';
import * as strings from 'ChartappWebPartStrings';

import { SPComponentLoader } from "@microsoft/sp-loader";

import "jquery";

import { sp } from "@pnp/sp";

import '../../CustomLogic/style.css'
import { func } from 'prop-types';

const d3: any = require("d3");

declare var $;

export interface IChartappWebPartProps {
  description: string;
}

export default class ChartappWebPart extends BaseClientSideWebPart<IChartappWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `<div class="widget">
    <div class="header"><h4>Project Status</h4></div>
    <div id="chart"></div>
    <div id="mainchartlegend"></div>
    <div class="commonlegend" id="alltaskcommonlegend"></div>
    <hr class="margin-top-40">
    <div class="margin-top-40">
    <input type="radio" class="clsmonth" name="month" value="3" checked>3 months<br>
    <input type="radio" class="clsmonth" name="month" value="6">6 months<br>   
    <div id="monthchart"></div>
    <div id="monthchartlegend"></div>
    </div>
    <hr class="margin-top-40">
    <div class="margin-top-40">
    <div id="userchart" style="margin-top: 45px;"></div>
    <div id="userchartlegend"></div>
    <div class="commonlegend" id="usertaskcommonlegend"></div>
    </div>
   `;


    var data = [];
    var allTasks = [];
    var userTasks = [];
    var userCompletedTasks = 0;
    var userSignificantDelayTasks = 0;
    var userSlightDelayTasks = 0;
    var userScheduleTasks = 0;
    var userIschild = false;
    var userTotalCount = 0;
    var userpercentage = {
      completed: 0, //blue
      overDue: 0, //red
      due: 0, //yellow
      upcoming: 0
    };

    var totalTasks = 0;
    var completedTasks = 0;
    var significantDelayTasks = 0;
    var slightDelayTasks = 0;
    var scheduleTasks = 0;
    var percentage = {
      completed: 0, //blue
      overDue: 0, //red
      due: 0, //yellow
      upcoming: 0
    };
    var ischild = false;
    var totalCount = 0;
    var datastatus = {
      completed: 'Completed', //blue
      overDue: 'Significant delay', //red
      due: 'Slight delay', //yellow
      upcoming: 'Ahead of schedule' //green
    };

    var currentUser = '';


    //Get current user details
    sp.web.currentUser.get().then(userdata => {

      currentUser = userdata.Id + '';

      //Fetching all data from the list
      sp.web.lists.getByTitle("PerformanceTracker19").items.getAll().then((allItems: any[]) => {

        allTasks = allItems;

        var currentDate = new Date();
        var beforeSixMonth = new Date();
        beforeSixMonth.setMonth(beforeSixMonth.getMonth() - 6);
        var overDue = new Date(beforeSixMonth.setDate(beforeSixMonth.getDate() - 1));

        totalTasks = allItems.length;

        //Seperating the projects based on the status
        for (var j = 0; j < allItems.length; j++) {
          var dueDate = new Date(allItems[j].DueDate);

          //Get completed task
          if (allItems[j].Status == datastatus.completed) {

            completedTasks = completedTasks + 1;
            var filterdata = findIndex(data, datastatus.completed);
            if (filterdata >= 0) {
              data[filterdata].children.push(
                {
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                });
            } else {
              data.push({
                name: datastatus.completed,
                color: '#0066ff',//blue,
                children: [{
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                }]
              });
            }

            var checkuserproperty = 'Executive_x0020_LeadStringId';

            //Get completed task for current user
            if (allItems[j][checkuserproperty] == currentUser) {
              userCompletedTasks = userCompletedTasks + 1;
              var filterdata = findIndex(userTasks, datastatus.completed);
              if (filterdata >= 0) {
                userTasks[filterdata].children.push(
                  {
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  });
              } else {
                userTasks.push({
                  name: datastatus.completed,
                  color: '#0066ff',//blue,
                  children: [{
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  }]
                });
              }
            }


          }

          //Get significant delay task
          else if (dueDate < currentDate && dueDate <= overDue) {

            significantDelayTasks = significantDelayTasks + 1;
            var filterdata = findIndex(data, datastatus.overDue);
            if (filterdata >= 0) {
              data[filterdata].children.push(
                {
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                });
            } else {
              data.push({
                name: datastatus.overDue,
                color: 'red',//red
                children: [{
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                }]
              });
            }

            //Get significant delay task for current user
            if (allItems[j][checkuserproperty] == currentUser) {
              userSignificantDelayTasks = userSignificantDelayTasks + 1;
              var filterdata = findIndex(userTasks, datastatus.overDue);
              if (filterdata >= 0) {
                userTasks[filterdata].children.push(
                  {
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  });
              } else {
                userTasks.push({
                  name: datastatus.overDue,
                  color: 'red',//red
                  children: [{
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  }]
                });
              }
            }
          }

          //Get slight delay task
          else if (dueDate < currentDate && dueDate >= beforeSixMonth) {
            slightDelayTasks = slightDelayTasks + 1;
            var filterdata = findIndex(data, datastatus.due);
            if (filterdata >= 0) {
              data[filterdata].children.push(
                {
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                });
            } else {
              data.push({
                name: datastatus.due,
                color: '#ffff1a',//yellow
                children: [{
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                }]
              });
            }


            //Get slight delay task for current user
            if (allItems[j][checkuserproperty] == currentUser) {
              userSlightDelayTasks = userSlightDelayTasks + 1;
              var filterdata = findIndex(userTasks, datastatus.due);
              if (filterdata >= 0) {
                userTasks[filterdata].children.push(
                  {
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  });
              } else {
                userTasks.push({
                  name: datastatus.due,
                  color: '#ffff1a',//yellow
                  children: [{
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  }]
                });
              }
            }

          }

          //Get schedule task
          else if (dueDate >= currentDate) {
            scheduleTasks = scheduleTasks + 1;
            var filterdata = findIndex(data, datastatus.upcoming);
            if (filterdata >= 0) {
              data[filterdata].children.push(
                {
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                });
            } else {
              data.push({
                name: datastatus.upcoming,
                color: '#00ff00',//green
                children: [{
                  name: allItems[j]["Title"],
                  percent: 10,
                  color: getRandomColor()
                }]
              });
            }

            //Get schedule task for current user
            if (allItems[j][checkuserproperty] == currentUser) {
              userScheduleTasks = userScheduleTasks + 1;
              var filterdata = findIndex(userTasks, datastatus.upcoming);
              if (filterdata >= 0) {
                userTasks[filterdata].children.push(
                  {
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  });
              } else {
                userTasks.push({
                  name: datastatus.upcoming,
                  color: '#00ff00',//green
                  children: [{
                    name: allItems[j]["Title"],
                    percent: 10,
                    color: getRandomColor()
                  }]
                });
              }
            }

          }
        }

        //Calculating percentage for all tasks
        var filterdata = findIndex(data, datastatus.completed);
        if (filterdata >= 0) {
          data[filterdata].percent = data[filterdata].children ? data[filterdata].children.length : 1;
        }
        filterdata = findIndex(data, datastatus.overDue);
        if (filterdata >= 0) {
          data[filterdata].percent = data[filterdata].children ? data[filterdata].children.length : 1;
        }
        filterdata = findIndex(data, datastatus.due);
        if (filterdata >= 0) {
          data[filterdata].percent = data[filterdata].children ? data[filterdata].children.length : 1;
        }
        filterdata = findIndex(data, datastatus.upcoming);
        if (filterdata >= 0) {
          data[filterdata].percent = data[filterdata].children ? data[filterdata].children.length : 1;
        }

        percentage.completed = (completedTasks / totalTasks) * 100;
        percentage.due = (slightDelayTasks / totalTasks) * 100;
        percentage.overDue = (significantDelayTasks / totalTasks) * 100;
        percentage.upcoming = (scheduleTasks / totalTasks) * 100;



        //Calculating percentage for current user tasks
        var filterdata = findIndex(userTasks, datastatus.completed);
        if (filterdata >= 0) {
          userTasks[filterdata].percent = userTasks[filterdata].children ? userTasks[filterdata].children.length : 1;
        }
        filterdata = findIndex(userTasks, datastatus.overDue);
        if (filterdata >= 0) {
          userTasks[filterdata].percent = userTasks[filterdata].children ? userTasks[filterdata].children.length : 1;
        }
        filterdata = findIndex(userTasks, datastatus.due);
        if (filterdata >= 0) {
          userTasks[filterdata].percent = userTasks[filterdata].children ? userTasks[filterdata].children.length : 1;
        }
        filterdata = findIndex(userTasks, datastatus.upcoming);
        if (filterdata >= 0) {
          userTasks[filterdata].percent = userTasks[filterdata].children ? userTasks[filterdata].children.length : 1;
        }

        userpercentage.completed = (userCompletedTasks / userTasks.length) * 100;
        userpercentage.due = (userSlightDelayTasks / userTasks.length) * 100;
        userpercentage.overDue = (userSignificantDelayTasks / userTasks.length) * 100;
        userpercentage.upcoming = (userScheduleTasks / userTasks.length) * 100;

        //Load all user chart
        updateGraph(data);

        //Load filter chart
        loadFilter();

        //Load chart for current user
        loadUserTask(userTasks);

        $('#alltaskcommonlegend').css('border', 'none');
        $('#alltaskcommonlegend').css('border', '1px solid');
        $("#alltaskcommonlegend").append('<div class="square" style="background: #0066ff;"></div>Completed<br/>');
        $("#alltaskcommonlegend").append('<div class="square" style="background: red;"></div>Significant Delay<br/>');
        $("#alltaskcommonlegend").append('<div class="square" style="background: #ffff1a;"></div>Slight Delay<br/>');
        $("#alltaskcommonlegend").append('<div class="square" style="background: #00ff00;"></div>Upcoming<br/>');


        $('#usertaskcommonlegend').css('border', 'none');
        $('#usertaskcommonlegend').css('border', '1px solid');
        $("#usertaskcommonlegend").append('<div class="square" style="background: #0066ff;"></div>Completed<br/>');
        $("#usertaskcommonlegend").append('<div class="square" style="background: red;"></div>Significant Delay<br/>');
        $("#usertaskcommonlegend").append('<div class="square" style="background: #ffff1a;"></div>Slight Delay<br/>');
        $("#usertaskcommonlegend").append('<div class="square" style="background: #00ff00;"></div>Upcoming<br/>');




        $(document).on('change', '.clsmonth', function () {
          loadFilter();
        });

      });
    });


    //Helper method to get chart data
    function findIndex(data, status) {
      return data.findIndex(c => c.name == status);
    }

    //Generate random color 
    function getRandomColor() {
      var letters = '0123456789ABCDEF';
      var color = '#';
      for (var i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
      }
      return color;
    }

    //Filter chart part
    function loadFilter() {
      try {


        var selectedFilter = $("input[name='month']:checked").val();
        var datavalue = [];
        var currentdate = new Date();
        var date = new Date();
        var filterDate;
        if (selectedFilter == '3') {
          filterDate = new Date(date.setMonth(date.getMonth() + 3));
        } else {
          filterDate = new Date(date.setMonth(date.getMonth() + 6));
        }
        for (let index = 0; index < allTasks.length; index++) {
          const element = allTasks[index];
          var dueDate = new Date(element.DueDate);
          if (dueDate >= currentdate && dueDate <= filterDate) {
            if (allTasks[index].Status != datastatus.completed) {
              datavalue.push({
                name: element.Title,
                percent: 10,
                color: getRandomColor()
              });
            }
          }
        }

        if (datavalue.length > 0) {
          $('#monthchart').empty();
          var pie = d3.layout.pie()
            .value(function (d) { return d.percent })
            .sort(null)
            .padAngle(.03);
          var w = 300, h = 300;
          var outerRadius = w / 2;
          var innerRadius = 100;
          var color = d3.scale.category10();
          var arc = d3.svg.arc()
            .outerRadius(outerRadius)
            .innerRadius(innerRadius);
          var svg = d3.select("#monthchart")
            .append("svg")
            .attr({
              width: w,
              height: h,
              class: 'shadow'
            }).append('g')
            .attr({
              transform: 'translate(' + w / 2 + ',' + h / 2 + ')'
            });
          svg.selectAll('path')
            .data(pie(datavalue))
            .enter()
            .append('path')
            .attr({
              d: arc,
              fill: function (d, i) {
                return color(d.data.name);
              }
            }).style("fill", function (d) {
              return d.data.color;
            });

          var g = svg.selectAll(".arc")
            .data(pie(datavalue))
            .enter().append("g");
          g.append("text")
            .attr("text-anchor", "middle")
            .attr('font-size', '4em')
            .attr('y', 20)
            .text(datavalue.length);
          $("#monthchartlegend").empty();
          $('#monthchartlegend').css('border', 'none');
          $('#monthchartlegend').css('border', '1px solid');
          for (let index = 0; index < datavalue.length; index++) {
            const element = datavalue[index];
            var style = 'style="background: ' + element.color + ';"';
            $("#monthchartlegend").append('<div class="square" ' + style + '></div>' + element.name + '<br/>');
          }
        }

        if (datavalue.length == 0) {
          $('#monthchart').empty();
          $('#monthchart').css('margin-top', '10px');
          $('#monthchart').append('<h3>No tasks to view</h3>');
        }
      }
      catch (err) {
        console.log(err.message);
      }
    }


    //Load current user chart
    function loadUserTask(datavalue) {
      try {
        $('#userchart').empty();
        var pie = d3.layout.pie()
          .value(function (d) { return d.percent })
          .sort(null)
          .padAngle(.03);

        var w = 300, h = 300;

        var outerRadius = w / 2;
        var innerRadius = 100;

        var color = d3.scale.category10();

        var arc = d3.svg.arc()
          .outerRadius(outerRadius)
          .innerRadius(innerRadius);

        var svg = d3.select("#userchart")
          .append("svg")
          .attr({
            width: w,
            height: h,
            class: 'shadow'
          }).append('g')
          .attr({
            transform: 'translate(' + w / 2 + ',' + h / 2 + ')'
          });

        var path = svg.selectAll('path')
          .data(pie(datavalue))
          .enter()
          .append('path')
          .attr({
            d: arc,
            fill: function (d, i) {
              return color(d.data.name);
            }
          }).style("fill", function (d) {
            return d.data.color;
          });


        var tooltip = d3.select('#userchart')
          .append('div')
          .attr('class', 'tooltip-margin');

        tooltip.append('div')
          .attr('class', 'label');

        $("#userchartlegend").empty();
        $('#userchartlegend').css('border', 'none');

        if (userIschild) {
          $('#usertaskcommonlegend').hide();

          var g = svg.selectAll(".arc")
            .data(pie(userTasks))
            .enter().append("g");
          g.append("text")
            .attr("text-anchor", "middle")
            .attr('font-size', '4em')
            .attr('y', 20)
            .text(userTotalCount);

          $('#userchartlegend').css('border', '1px solid');
          for (let index = 0; index < datavalue.length; index++) {
            const element = datavalue[index];
            var style = 'style="background: ' + element.color + ';"';
            $("#userchartlegend").append('<div class="square" ' + style + '></div>' + element.name + '<br/>');
          }

        } else {
          $('#usertaskcommonlegend').show();
          var g = svg.selectAll(".arc")
            .data(pie(datavalue))
            .enter().append("g");

          g.append("text")
            .attr("text-anchor", "middle")
            .attr('font-size', '4em')
            .attr('y', 20)
            .text(userTasks.length);

          g.append("text")
            .attr("transform", function (d) {
              var _d = arc.centroid(d);
              _d[0] *= 1.5;	//multiply by a constant factor
              _d[1] *= 1.4;	//multiply by a constant factor
              return "translate(" + _d + ")";
            })
            .attr("dy", ".50em")
            .style("text-anchor", "middle")
            .text(function (d) {
              if (d.data.name == datastatus.completed) {
                return userpercentage.completed.toFixed(2) + ' %';
              } else if (d.data.name == datastatus.due) {
                return userpercentage.due.toFixed(2) + ' %';
              } else if (d.data.name == datastatus.overDue) {
                return userpercentage.overDue.toFixed(2) + ' %';
              } else if (d.data.name == datastatus.upcoming) {
                return userpercentage.upcoming.toFixed(2) + ' %';
              }
            });
        }

        path.on("click", function (d) {
          if (d.data.children) {
            var textlabel = d.data.children.length + ' ' + d.data.name;
            userIschild = true;
            userTotalCount = d.data.children.length;
          } else {
            userIschild = false;
            $('#data').html('')
          }
          loadUserTask(d.data.children ? d.data.children : userTasks);
        })
          .on("mouseover", function (d) {
            var ang = d.startAngle + (d.endAngle - d.startAngle) / 2;
            ang = (ang - (Math.PI / 2)) * -1;
            var width = 500,
              height = 400,
              margin = 50,
              radius = Math.min(width - margin, height - margin) / 2;
            var x = Math.cos(ang) * radius * 0.1;
            var y = Math.sin(ang) * radius * -0.1;
            d3.select(this).transition()
              .duration(250).attr("transform", "translate(" + x + "," + y + ")");

            tooltip.select('.label').html(d.data.name);
            tooltip.style('display', 'block');

          })
          .on("mouseout", function (d) {
            d3.select(this).transition()
              .duration(150).attr("transform", "translate(0,0)");
            tooltip.style('display', 'none');
          });

        if (datavalue.length == 0) {
          $('#userchart').empty();
          $('#userchart').css('margin-top', '10px');
          $('#userchart').append('<h3>No tasks are assigned to the user</h3>');
        }
      }
      catch (err) {
        console.log(err.message);
      }
    }

    //Load all data into chart
    function updateGraph(datavalue) {
      try {
        $('#mainchartlegend').empty();
        $('#mainchartlegend').css('border', 'none');

        $('#chart').empty();
        var pie = d3.layout.pie()
          .value(function (d) { return d.percent })
          .sort(null)
          .padAngle(.03);

        var w = 300, h = 300;

        var outerRadius = w / 2;
        var innerRadius = 100;

        var color = d3.scale.category10();

        var arc = d3.svg.arc()
          .outerRadius(outerRadius)
          .innerRadius(innerRadius);

        var svg = d3.select("#chart")
          .append("svg")
          .attr({
            width: w,
            height: h,
            class: 'shadow'
          }).append('g')
          .attr({
            transform: 'translate(' + w / 2 + ',' + h / 2 + ')'
          });

        var path = svg.selectAll('path')
          .data(pie(datavalue))
          .enter()
          .append('path')
          .attr({
            d: arc,
            fill: function (d, i) {
              return color(d.data.name);
            }
          }).style("fill", function (d) {
            return d.data.color;
          });


        var tooltip = d3.select('#chart')
          .append('div')
          .attr('class', 'tooltip');

        tooltip.append('div')
          .attr('class', 'label');


        var div = d3.select("#data").append("div")
          .attr("class", "tooltip-donut")
          .style("opacity", 0);



        if (ischild) {
          $('#alltaskcommonlegend').hide();
          var g = svg.selectAll(".arc")
            .data(pie(datavalue))
            .enter().append("g");
          g.append("text")
            .attr("text-anchor", "middle")
            .attr('font-size', '4em')
            .attr('y', 20)
            .text(totalCount);

          $('#mainchartlegend').css('border', '1px solid');
          for (let index = 0; index < datavalue.length; index++) {
            const element = datavalue[index];
            var style = 'style="background: ' + element.color + ';"';
            $("#mainchartlegend").append('<div class="square" ' + style + '></div>' + element.name + '<br/>');
          }

        } else {
          $('#alltaskcommonlegend').show();
          $("#mainchartlegend").empty();

          var g = svg.selectAll(".arc")
            .data(pie(datavalue))
            .enter().append("g");

          g.append("text")
            .attr("text-anchor", "middle")
            .attr('font-size', '4em')
            .attr('y', 20)
            .text(totalTasks);

          g.append("text")
            .attr("transform", function (d) {
              var _d = arc.centroid(d);
              _d[0] *= 1.5;	//multiply by a constant factor
              _d[1] *= 1.4;	//multiply by a constant factor
              return "translate(" + _d + ")";
            })
            .attr("dy", ".50em")
            .style("text-anchor", "middle")
            .text(function (d) {
              if (d.data.name == datastatus.completed) {
                return percentage.completed.toFixed(2) + ' %';
              } else if (d.data.name == datastatus.due) {
                return percentage.due.toFixed(2) + ' %';
              } else if (d.data.name == datastatus.overDue) {
                return percentage.overDue.toFixed(2) + ' %';
              } else if (d.data.name == datastatus.upcoming) {
                return percentage.upcoming.toFixed(2) + ' %';
              }
            });
        }

        path.on("click", function (d) {
          if (d.data.children) {
            var textlabel = d.data.children.length + ' ' + d.data.name;
            ischild = true;
            totalCount = d.data.children.length;
            // $('#data').html('<h3>' + textlabel + '</h3>');
          } else {
            ischild = false;
            $('#data').html('')
          }
          updateGraph(d.data.children ? d.data.children : data);
        })
          .on("mouseover", function (d) {
            var ang = d.startAngle + (d.endAngle - d.startAngle) / 2;
            ang = (ang - (Math.PI / 2)) * -1;
            var width = 500,
              height = 400,
              margin = 50,
              radius = Math.min(width - margin, height - margin) / 2;
            var x = Math.cos(ang) * radius * 0.1;
            var y = Math.sin(ang) * radius * -0.1;
            d3.select(this).transition()
              .duration(250).attr("transform", "translate(" + x + "," + y + ")");

            tooltip.select('.label').html(d.data.name);
            tooltip.style('display', 'block');
          })
          .on("mouseout", function (d) {
            // div.html('');
            d3.select(this).transition()
              .duration(150).attr("transform", "translate(0,0)");

            tooltip.style('display', 'none');

          });

        if (datavalue.length == 0) {
          $('#chart').empty();
          $('#chart').css('margin-top', '10px');
          $('#chart').append('<h3>No tasks to view</h3>');
        }
      }
      catch (err) {
        console.log(err.message);
      }
    }

  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
