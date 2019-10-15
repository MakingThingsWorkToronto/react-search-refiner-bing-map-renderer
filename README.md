# react-search-refiner-bing-map-renderer

![Version](https://img.shields.io/badge/version-0.0.1-green.svg)

## Summary
The Bing Maps Search Renderer is a custom code renderer for react-search-refiners project located at: https://github.com/SharePoint/sp-dev-solutions/tree/master/solutions/ModernSearch/react-search-refiners

This solution contains two web parts that display search results on a map:

1) bingMapsPolygonSearch: Draws a compressed polygon shape on a bing map.
- Draws polygons compressed with [PointCompression](https://docs.microsoft.com/en-us/bingmaps/v8-web-control/map-control-api/pointcompression-class).
- Customizable colors.
- Custom hyperlinks.


2) bingMapsSearch: Plots points on a map containing templated infoboxes & customizable pins.
- Plots points on map by mapping to TWO text search managed properties: one containing latitude, one containing longitude.
- Most Bing Maps options supported.
- Apply custom pin images via regular expressions with multi value property compatability.
- Create custom info box pop-ups in the browser using Handlerbarsjs templates.


### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.
* src/* - all source code, web parts, models and services

### Build options

gulp clean
gulp test - TODO
gulp serve
gulp bundle
gulp package-solution
