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
// import "canvas";

import { sp } from "@pnp/sp";

var d3layout: any = require("../../CustomLogic/d3.layout.js");

const d3: any = require("d3");

declare var $;



export interface IChartappWebPartProps {
  description: string;
}

export default class ChartappWebPart extends BaseClientSideWebPart<IChartappWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `<h2>
    Project status
  </h2><div id="data"></div>`;

    // Globals
    var width = 500,
      height = 400,
      margin = 50,
      radius = Math.min(width - margin, height - margin) / 2,
      // Pie layout will use the "val" property of each data object entry
      pieChart = d3.layout.pie().sort(null).value(function (d) { return d.val; }),
      arc = d3.svg.arc().outerRadius(radius),
      MAX_SECTORS = 15, // Less than 20 please
      colors = d3.scale.category20();

    var data = [];

    sp.web.lists.getByTitle("Project Tasks").items.getAll().then((allItems: any[]) => {
      var currentDate = new Date();
      var beforeSixMonth = new Date();
      beforeSixMonth.setMonth(beforeSixMonth.getMonth() - 6);
      var overDue = new Date(beforeSixMonth.setDate(beforeSixMonth.getDate() - 1));

      var values = {
        compeleted: 0.6012300409454765,
        overDue: 0.6012300409454765,
        due: 0.6012300409454765,
        upcoming: 0.6012300409454765,
      };

      for (var j = 0; j < allItems.length; j++) {
        var dueDate = new Date(allItems[j].DueDate);

        if (allItems[j].Status == 'Compeleted') {
          var filterdata = findIndex(data, 'Compeleted');
          if (filterdata >= 0) {
            data[filterdata].children.push(
              {
                cat: allItems[j]["Title"],
                val: 10,
                color: getRandomColor()
              });
          } else {
            data.push({
              cat: 'Compeleted',
              val: 0.6012300409454765,
              color: '#0066ff',//blue,
              children: [{
                cat: allItems[j]["Title"],
                val: 10,
                color: getRandomColor()
              }]
            });
          }
          values.compeleted = values.compeleted * 2;
        }

        else if (dueDate < currentDate) {
          if (dueDate <= overDue) {
            var filterdata = findIndex(data, 'Over due');
            if (filterdata >= 0) {
              data[filterdata].children.push(
                {
                  cat: allItems[j]["Title"],
                  val: 10,
                  color: getRandomColor()
                });
            } else {
              data.push({
                cat: 'Over due',
                val: 0.6012300409454765,
                color: '#d147a3',//red
                children: [{
                  cat: allItems[j]["Title"],
                  val: 10,
                  color: getRandomColor()
                }]
              });
            }
            values.overDue = values.overDue * 2;
          }
        }

        if (dueDate < currentDate) {
          if (dueDate >= beforeSixMonth) {
            var filterdata = findIndex(data, 'Due');
            if (filterdata >= 0) {
              data[filterdata].children.push(
                {
                  cat: allItems[j]["Title"],
                  val: 10,
                  color: getRandomColor()
                });
            } else {
              data.push({
                cat: 'Due',
                val: 0.6012300409454765,
                color: '#ffff1a',//yellow
                children: [{
                  cat: allItems[j]["Title"],
                  val: 10,
                  color: getRandomColor()
                }]
              });
            }
            values.due = values.due * 2;
          }
        }

        else if (dueDate >= currentDate) {
          var filterdata = findIndex(data, 'Upcoming');
          if (filterdata >= 0) {
            data[filterdata].children.push(
              {
                cat: allItems[j]["Title"],
                val: 10,
                color: getRandomColor()
              });
          } else {
            data.push({
              cat: 'Upcoming',
              val: 0.6012300409454765,
              color: '#00ff00',//green
              children: [{
                cat: allItems[j]["Title"],
                val: 10,
                color: getRandomColor()
              }]
            });
          }
          values.upcoming = values.upcoming * 2;
        }
      }
      var filterdata = findIndex(data, 'Compeleted');
      if (filterdata >= 0) {
        data[filterdata].val = values.compeleted;
      }
      filterdata = findIndex(data, 'Over due');
      if (filterdata >= 0) {
        data[filterdata].val = values.overDue;
      }
      filterdata = findIndex(data, 'Due');
      if (filterdata >= 0) {
        data[filterdata].val = values.due;
      }
      filterdata = findIndex(data, 'Upcoming');
      if (filterdata >= 0) {
        data[filterdata].val = values.upcoming;
      }

      // Start by updating graph at root level
      updateGraph(undefined);

    });

    function findIndex(data, status) {
      return data.findIndex(c => c.cat == status);
    }

    function getRandomColor() {
      var letters = '0123456789ABCDEF';
      var color = '#';
      for (var i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
      }
      return color;
    }


    // var numSectors = Math.ceil(Math.random() * MAX_SECTORS);
    // for (let i = -1; i++ < numSectors;) {
    //   var children = [];
    //   var numChildSectors = Math.ceil(Math.random() * MAX_SECTORS);
    //   var color = colors(i);
    //   for (let j = -1; j++ < numChildSectors;) {
    //     // Add children categories with shades of the parent color
    //     children.push(
    //       {
    //         cat: "cat" + ((i + 1) * 100 + j),
    //         val: Math.random(),
    //         color: colors((j + 1) * 10) //d3.rgb(color).darker(1 / (j + 1))
    //       });
    //   }
    //   data.push({
    //     cat: "cat" + i,
    //     val: Math.random(),
    //     color: color,
    //     children: children
    //   });
    // }
    // --------------------------------------------------------------------------


    // SVG elements init
    var svg = d3.select("#data").append("svg").attr("width", width).attr("height", height),
      defs = svg.append("svg:defs"),
      // Declare a main gradient with the dimensions for all gradient entries to refer
      mainGrad = defs.append("svg:radialGradient")
        .attr("gradientUnits", "userSpaceOnUse")
        .attr("cx", 0).attr("cy", 0).attr("r", radius).attr("fx", 0).attr("fy", 0)
        .attr("id", "master"),
      // The pie sectors container
      arcGroup = svg.append("svg:g")
        .attr("class", "arcGroup")
        .attr("filter", "url(#shadow)")
        .attr("transform", "translate(" + (width / 2) + "," + (height / 2) + ")"),
      // Header text
      header = svg.append("text").text("Parent")
        .attr("transform", "translate(10, 20)").attr("class", "header");

    // Declare shadow filter
    var shadow = defs.append("filter").attr("id", "shadow")
      .attr("filterUnits", "userSpaceOnUse")
      .attr("x", -1 * (width / 2)).attr("y", -1 * (height / 2))
      .attr("width", width).attr("height", height);
    shadow.append("feGaussianBlur")
      .attr("in", "SourceAlpha")
      .attr("stdDeviation", "4")
      .attr("result", "blur");
    shadow.append("feOffset")
      .attr("in", "blur")
      .attr("dx", "4").attr("dy", "4")
      .attr("result", "offsetBlur");
    shadow.append("feBlend")
      .attr("in", "SourceGraphic")
      .attr("in2", "offsetBlur")
      .attr("mode", "normal");



    // Redraw the graph given a certain level of data
    function updateGraph(cat) {
      var currData = data;

      // Simple header text
      if (cat != undefined) {
        currData = findChildenByCat(cat);
        d3.select(".header").text("Parent → " + cat);
      } else {
        d3.select(".header").text("Parent");
      }

      // Create a gradient for each entry (each entry identified by its unique category)
      var gradients = defs.selectAll(".gradient").data(currData, function (d) { return d.cat; });
      gradients.enter().append("svg:radialGradient")
        .attr("id", function (d, i) { return "gradient" + d.cat; })
        .attr("class", "gradient")
        .attr("xlink:href", "#master");

      gradients.append("svg:stop").attr("offset", "0%").attr("stop-color", getColor);
      gradients.append("svg:stop").attr("offset", "90%").attr("stop-color", getColor);
      gradients.append("svg:stop").attr("offset", "100%").attr("stop-color", getDarkerColor);


      // Create a sector for each entry in the enter selection
      var paths = arcGroup.selectAll("path")
        .data(pieChart(currData), function (d) { return d.data.cat; });
      paths.enter().append("svg:path").attr("class", "sector");

      // Each sector will refer to its gradient fill
      paths.attr("fill", function (d, i) { return "url(#gradient" + d.data.cat + ")"; })
        .transition().duration(1000).attrTween("d", tweenIn).each("end", function () {
          this._listenToEvents = true;
        });


      var div = d3.select("#data").append("div")
        .attr("class", "tooltip-donut")
        .style("opacity", 0);


      // Mouse interaction handling
      paths.on("click", function (d) {
        if (this._listenToEvents) {
          // Reset inmediatelly
          d3.select(this).attr("transform", "translate(0,0)")
          // Change level on click if no transition has started                
          paths.each(function () {
            this._listenToEvents = false;
          });

          // if (d.data.cat) {
          //   if (d.data.children && d.data.children.length > 0) {
          //     updateGraph(d.data.children);
          //   } else {
          //     updateGraph(undefined);
          //   }
          // } else {
          //   updateGraph(d.data.children ? d.data.cat : undefined);
          // }
          updateGraph(d.data.children ? d.data.cat : undefined);
        }
      })
        .on("mouseover", function (d) {
          // Mouseover effect if no transition has started                
          if (this._listenToEvents) {
            // Calculate angle bisector
            var ang = d.startAngle + (d.endAngle - d.startAngle) / 2;
            // Transformate to SVG space
            ang = (ang - (Math.PI / 2)) * -1;

            // Calculate a 10% radius displacement
            var x = Math.cos(ang) * radius * 0.1;
            var y = Math.sin(ang) * radius * -0.1;

            d3.select(this).transition()
              .duration(250).attr("transform", "translate(" + x + "," + y + ")");

            div.transition()
              .duration(50)
              .style("opacity", 1);

            $('.tooltip-donut').empty();

            div.html(d.data.cat)
              .style("left", (d3.event.pageX + 10) + "px")
              .style("top", (d3.event.pageY - 15) + "px");
          }
        })
        .on("mouseout", function (d) {
          // Mouseout effect if no transition has started                
          if (this._listenToEvents) {
            d3.select(this).transition()
              .duration(150).attr("transform", "translate(0,0)");
          }
        });

      // Collapse sectors for the exit selection
      paths.exit().transition()
        .duration(1000)
        .attrTween("d", tweenOut).remove();
    }


    // "Fold" pie sectors by tweening its current start/end angles
    // into 2*PI
    function tweenOut(data) {
      data.startAngle = data.endAngle = (2 * Math.PI);
      var interpolation = d3.interpolate(this._current, data);
      this._current = interpolation(0);
      return function (t) {
        return arc(interpolation(t));
      };
    }


    // "Unfold" pie sectors by tweening its start/end angles
    // from 0 into their final calculated values
    function tweenIn(data) {
      var interpolation = d3.interpolate({ startAngle: 0, endAngle: 0 }, data);
      this._current = interpolation(0);
      return function (t) {
        return arc(interpolation(t));
      };
    }


    // Helper function to extract color from data object
    function getColor(data, index) {
      return data.color;
    }


    // Helper function to extract a darker version of the color
    function getDarkerColor(data, index) {
      return d3.rgb(getColor(data, index)).darker();
    }


    function findChildenByCat(cat) {
      for (let i = -1; i++ < data.length - 1;) {
        if (data[i].cat == cat) {
          return data[i].children;
        }
      }
      return data;
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
