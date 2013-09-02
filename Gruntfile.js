module.exports = function(grunt) {

    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),

        // Tests, via Buster
        buster: {
            default: {
                test: {
                    config: 'test/buster.js',
                    'config-group': 'node'
                }
            }
        }
    });

    grunt.loadNpmTasks('grunt-buster');

    // Aliases
    grunt.registerTask('test', ['buster:default']);
};