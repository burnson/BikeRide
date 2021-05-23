'use strict';
/*jslint browser: true*/
/*globals download*/

function Bike(parent) {
  this.parent = parent !== undefined ? parent : document.body;
  this.apparentWidth = null;
  this.apparentHeight = null;
  this.actualWidth = null;
  this.actualHeight = null;
  this.canvas = null;
  this.showReference = true;

  // Bike dimensions
  this.canvasWidth = 32;
  this.canvasHeight = 19;

  // Bike frame constants
  this.spaceHeights = 0.02;
  this.lineThicknesses = this.spaceHeights * 0.22;
  this.bikeFrameRadius = 2.0;
  this.bikeFrameInnerWidth = 3.1;
  this.bikeFrameUpperFrontSegmentHeight = 0.3;
  this.bikeFrameScale = 2.8;
  this.bikeFrameUpperBackPosition = {x: -1.3, y: this.bikeFrameRadius * 2 + 1};
  this.bikeFramePedalOrigin = {x: -0.1, y: -0.7 + this.bikeFrameRadius};
  this.bikeFrameOrigin = {x: 32.0 / 2.0, y: 1.6};
  this.GroundWidth = 4.0;
  this.GroundYDisplacement = this.SpaceHeights * -10.5;
  this.GroundLineExtension = 5.5;
  this.LeftGearsArcWidth = 0.5;

  // Color constants
  this.frameColor = {r: 0.9, g: 0.9, b: 0.94};
  this.wheelColor = {r: 0.2, g: 0.3, b: 0.4};
  this.bikeGuardColor = {r: 0.2, g: 0.3, b: 0.4};
  this.leftOuterGearColor = {r: 0.2, g: 0.3, b: 0.4};
  this.leftInnerGearColor = {r: 0.2, g: 0.3, b: 0.4};
  this.rightGearColor = {r: 0.2, g: 0.3, b: 0.4};
  this.groundLineColor = {r: 0.2, g: 0.3, b: 0.4};
  this.groundColors = {r: 0.2, g: 0.3, b: 0.4};

  // Radii constants
  this.bikeGuardRadius = this.bikeFrameRadius + 0.35;
  this.leftBikeGearRadius = 0.4;
  this.leftBikeInnerGearRadius = 0.25;
  this.rightBikeGearRadius = 0.6;
  this.rightBikeInnerGearRadius = 0.45;
  this.leftSpokeRadius = 0.4 + this.spaceHeights * 2.0;
  this.rightSpokeRadius = 0.1;

  // Spoke constants
  this.spokeSlant = Math.PI / 16;
  this.spokesPerWheel = 40;
  this.spokeThickness = 0.003;

  // Staff point counts
  this.centerPoints = 200;
  this.leftBikeWheelPoints = 200;
  this.rightBikeWheelPoints = 200;
  this.leftBikeGuardPoints = 200;
  this.rightBikeGuardPoints = this.leftBikeGuardPoints / 2;
  this.leftBikeGearPoints = 400;
  this.leftBikeInnerGearPoints = 400;
  this.rightBikeGearPoints = 359;
  this.rightBikeInnerGearPoints = 200;
  this.groundPoints = 10;

  // Staff objects
  this.leftCenter = null;
  this.rightCenter = null;
  this.leftBikeWheel = null;
  this.rightBikeWheel = null;
  this.leftBikeGuard = null;
  this.rightBikeGuard = null;
  this.leftBikeGear = null;
  this.rightBikeGear = null;
  this.leftBikeInnerGear = null;
  this.rightBikeInnerGear = null;
  this.groundLeft = null;
  this.groundRight = null;
  this.groundLine = null;

  // Handlebar constants
  this.handleBarsHeight = 0.45;
  this.handleBarsExtension = 0.45;
  this.handleBarsRadius = 0.38;
  this.handleBarsCirclePercentage = 0.55;
  this.handleBarsCirclePoints = 200;

  // Seat constants
  this.seatHeight = 0.35;
  this.seatWidth = 0.75;
  this.seatTipAngle = Math.PI - Math.PI / 3;
  this.seatTipLength = 0.2;
  this.seatBackWidth = 0.2;

  // Landmarks
  this.point1 = {x: -this.bikeFrameInnerWidth, y: this.bikeFrameRadius};
  this.point2 = {x: this.bikeFrameInnerWidth, y: this.bikeFrameRadius};
  this.point3 = {
    x: this.bikeFramePedalOrigin.x,
    y: this.bikeFramePedalOrigin.y
  };
  this.point4 = {
    x: this.bikeFrameUpperBackPosition.x,
    y: this.bikeFrameUpperBackPosition.y
  };
  this.point5 = {
    x: this.point2.x + (this.point4.y - this.point2.y) /
      ((this.point4.y - this.point3.y) / (this.point4.x - this.point3.x)),
    y: this.point4.y
  };
  this.point6Connection = 0.91;
  this.point6 = {
    x: (this.point5.x - this.point2.x) * this.point6Connection + this.point2.x,
    y: (this.point5.y - this.point2.y) * this.point6Connection + this.point2.y
  };
  this.initialize();
}

Bike.prototype.getBackingStorePixelRatio = function (context) {
  return (context.webkitBackingStorePixelRatio ||
    context.mozBackingStorePixelRatio ||
    context.msBackingStorePixelRatio ||
    context.oBackingStorePixelRatio ||
    context.backingStorePixelRatio || 1);
};

Bike.prototype.getDevicePixelRatio = function () {
  return (window.devicePixelRatio || 1);
};

Bike.prototype.getHiDPIPixelRatio = function (context) {
  return this.getDevicePixelRatio() / this.getBackingStorePixelRatio(context);
};

Bike.prototype.drawPoint = function (point, color) {
  var context = this.canvas.getContext('2d');
  context.beginPath();
  context.ellipse(point.x, point.y, 0.1, 0.1, 0, 0, Math.PI * 2);
  context.fillStyle = color;
  context.fill();
};

Bike.prototype.drawLine = function (p1, p2, lineWidth, color, showPoints) {
  var context = this.canvas.getContext('2d');
  context.beginPath();
  context.moveTo(p1.x, p1.y);
  context.lineTo(p2.x, p2.y);
  context.lineWidth = lineWidth;
  context.lineCap = 'round';
  context.lineJoin = 'round';
  context.strokeStyle = color;
  context.stroke();
  if (showPoints) {
    this.drawPoint(p1, '#b4b0b0');
    this.drawPoint(p2, '#b4b0b0');
  }
};

Bike.prototype.drawGrid = function () {
  var i, w = this.canvasWidth, h = this.canvasHeight;
  for (i = 0; i <= w; i += 1) {
    this.drawLine({x: i, y: 0}, {x: i, y: h}, 0.05, '#f0f0f0');
  }
  for (i = 0; i <= h; i += 1) {
    this.drawLine({x: 0, y: i}, {x: w, y: i}, 0.05, '#f0f0f0');
  }
};

Bike.prototype.centerAndScaleCanvasInParent = function () {
  var context = this.canvas.getContext('2d'), xScale, yScale, scale;
  // context.translate(this.actualWidth / 2, this.actualHeight / 2);
  xScale = this.actualWidth / this.canvasWidth;
  yScale = this.actualHeight / this.canvasHeight;
  scale = xScale < yScale ? xScale : yScale;
  context.scale(scale, -scale);
  //context.translate(0, -(this.canvasHeight + this.actualHeight / scale) / 2);
  context.translate(0, -this.canvasHeight - 0.68);
  this.drawGrid();
};

Bike.prototype.draw = function () {
  var context = this.canvas.getContext('2d');
  this.centerAndScaleCanvasInParent();

  context.translate(16, 1.6);
  context.scale(2.8, 2.8);
  this.drawGrid();

  this.drawLine(this.point1, this.point4, 0.1, '#d9e3eb');
  this.drawLine(this.point4, this.point5, 0.1, '#d9e3eb');
  this.drawLine(this.point5, this.point2, 0.1, '#d9e3eb');
  this.drawLine(this.point4, this.point3, 0.1, '#d9e3eb');
  this.drawLine(this.point3, this.point6, 0.1, '#d9e3eb');
};

Bike.prototype.initialize = function () {
  var parentStyle = this.parent.style,
    referenceStyle = document.getElementById('reference').style,
    context = null,
    ratio = null,
    canvas = null;
  if (this.parent === document.body) {
    parentStyle.position = 'absolute';
    parentStyle.margin = '0px';
    parentStyle.left = '0px';
    parentStyle.right = '0px';
    parentStyle.top = '0px';
    parentStyle.bottom = '0px';
    parentStyle.overflow = 'hidden';
  }
  if (referenceStyle) {
    referenceStyle.position = 'absolute';
    referenceStyle.margin = '0px';
    referenceStyle.left = '0px';
    referenceStyle.right = '0px';
    referenceStyle.top = '0px';
    referenceStyle.bottom = '0px';
    referenceStyle.overflow = 'hidden';
    if (!this.showReference) {
      referenceStyle.opacity = 0;
    } else {
      referenceStyle.opacity = 0.1;
      referenceStyle.zIndex = 1;
    }
  }

  this.apparentWidth = this.parent.offsetWidth;
  this.apparentHeight = this.parent.offsetHeight;
  this.canvas = document.createElement('canvas');

  context = this.canvas.getContext('2d');
  ratio = this.getHiDPIPixelRatio(context);
  this.actualWidth = this.apparentWidth * ratio;
  this.actualHeight = this.apparentHeight * ratio;
  context.scale(ratio, ratio);

  this.canvas.width = this.actualWidth;
  this.canvas.height = this.actualHeight;
  this.canvas.style.width = this.apparentWidth + 'px';
  this.canvas.style.height = this.apparentHeight + 'px';

  this.parent.appendChild(this.canvas);

  canvas = this.canvas;
  this.canvas.addEventListener('click', function () {
    console.log('Downloading...');
    download(canvas.toDataURL(), 'image.png', 'image/png');
  });

  this.draw();
};

var bike = new Bike();
