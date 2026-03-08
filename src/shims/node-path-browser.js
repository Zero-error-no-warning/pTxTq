import {
  basenamePosixPath,
  dirnamePosixPath,
  extnamePosixPath,
  joinPosixPath,
  normalizePosixPath,
  relativePosixPath
} from "../utils/posixPath.js";

export const posix = {
  basename: basenamePosixPath,
  dirname: dirnamePosixPath,
  extname: extnamePosixPath,
  join: joinPosixPath,
  normalize: normalizePosixPath,
  relative: relativePosixPath
};

export default {
  posix,
  basename: basenamePosixPath,
  dirname: dirnamePosixPath,
  extname: extnamePosixPath,
  join: joinPosixPath,
  normalize: normalizePosixPath,
  relative: relativePosixPath
};
