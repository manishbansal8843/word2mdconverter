const path = require("path");
const gulp = require("gulp");
const exec = require("./scripts/exec");
const word2mdConverterJs = "./build/word2mdconverter.js";
const sampleMd = "./doc/word.md";
const typescript = require('gulp-tsc');
const clean = require('gulp-clean');

gulp.task("build", () =>
    exec("cscript", ["//nologo", word2mdConverterJs, path.resolve("./doc/word.docx"),path.resolve(sampleMd)]));


    gulp.task('clean', function () {
        return gulp.src('build', {read: false})
            .pipe(clean());
    });
    gulp.task('compile', function(){
      gulp.src(['src/*.ts'])
        .pipe(typescript())
        .pipe(gulp.dest('build/'))
    });