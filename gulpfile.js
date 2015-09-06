var gulp = require('gulp'),
    minifycss = require('gulp-minify-css'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename'),
    jshint = require('gulp-jshint');

gulp.task('minifycss', function () {
	return gulp.src('AppChromeControl-demo/css/Office.Controls.AppChrome.css')
		.pipe(rename({suffix:'.min'}))
		.pipe(gulp.dest('dist/'))
		.pipe(minifycss());
});

gulp.task('minifyjs', function () {
	return gulp.src(['AppChromeControl-demo/js/Office.Controls.AppChrome.js', 'AppChromeControl-demo/js/Office.Controls.Login.js'])
		.pipe(rename({suffix: '.min'}))
		.pipe(uglify({compress: true,mangle: true, outSourceMap: true}))
		.pipe(gulp.dest('dist/'));
});

gulp.task('cpfilestodist', ['minifycss', 'minifyjs'], function() {
    return gulp.src('src/**/*')
    .pipe(gulp.dest('dist/'));
});

gulp.task('cpfilestoexample', ['cpfilestodist'], function() {
    return gulp.src('dist/**/*')
    .pipe(gulp.dest('example/control/'));
});

gulp.task('default', ['cpfilestoexample']);