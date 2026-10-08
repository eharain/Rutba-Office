// 3D models: a glTF read into triangles (gltf.js) and drawn into pixels by
// the suite's own renderer (raster3d.js), for Insert → 3D Models.
export { readModel, parseGlb, isGlb, toGlb, gltfFiles } from './gltf.js';
export { renderModel, fit, outline, viewMatrix, MODEL_VIEWS } from './raster3d.js';
