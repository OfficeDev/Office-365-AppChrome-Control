var gulp = require('gulp'),
    minifycss = require('gulp-minify-css'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename'),
    jshint = require('gulp-jshint');

gulp.task('minifycss', function () {
	return gulp.src('src/Office.Controls.AppChrome.css')
		.pipe(rename({suffix:'.min'}))
		.pipe(minifycss())
		.pipe(gulp.dest('dist/'));
});

gulp.task('minifyjs', function () {
    return ['src/Office.Controls.AppChrome.js',
            'src/Office.Controls.Login.js'
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(rename({ suffix: '.min' }))
                    .pipe(uglify({compress: true,mangle: true, outSourceMap: true}))
                    .pipe(gulp.dest('dist/'));
                });
});

gulp.task('runjshint', function () {
    return ['src/Office.Controls.AppChrome.js',
            'src/Office.Controls.Login.js'
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(jshint('tools/jshint/.jshintrc.json'))
                    .pipe(jshint.reporter('jshint-stylish'));
                });
});

gulp.task('cpfilestodist', ['minifycss', 'minifyjs'], function() {
    return gulp.src('src/**/*')
    .pipe(gulp.dest('dist/'));
})

gulp.task('default', ['runjshint', 'cpfilestodist']);