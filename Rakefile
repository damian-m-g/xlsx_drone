desc 'compile & run the unit tests'
task :ct do
  build_unit_tests()
  run_unit_tests()
end

desc 'compile unit tests'
task :c do
  build_unit_tests()
end

desc 'run the unit tests'
task :t do
  run_unit_tests()
end

def build_unit_tests
  command = []
  command << gcc = 'gcc'
  command << debug = '-ggdb' # leave empty if the target exe hasn't the purpose of debugging
  command << standard = '-std=c11'
  command << warnings = '-w' # libraries used throw some warnings
  command << directories_included_for_headers_search = '-I ext/ -I src/'
  command << source_files = 'ext/zip.c ext/sxmlc.c ext/sxmlsearch.c src/library.c ext/unity.c src/library.test.c'
  command << output = '-o bin/unit_tests.exe'
  system(command.join(' '))
end

def run_unit_tests
  system("bin/unit_tests")
end