=begin
TASKS
=end

desc 'compile(x86) & run the unit tests'
task :ct do
  if(build_unit_tests())
    run_unit_tests()
  end
end

desc 'compile(x64) & run the unit tests'
task :ct_64 do
  if(build_unit_tests(_64: true))
    run_unit_tests()
  end
end

desc 'compile(x86) unit tests'
task :c do
  build_unit_tests()
end

desc 'compile(x64) unit tests'
task :c_64 do
  build_unit_tests(_64: true)
end

desc 'run the unit tests'
task :t do
  run_unit_tests()
end

desc 'clear unit tests'
task :clear_ut do
  clear_unit_tests()
end

# you will execute this before every new version release
desc 'perform measures & produce badges metadata'
task :badges do
  # not important if compiled with x86 or x64
  if(build_unit_tests(coverage: true))
    # if tests don't pass, code coverage will lie since code next to the problem isn't executed. In that case it's better to leave the last measure
    if(run_unit_tests())
      measure_code_coverage()
      create_test_suite_badge(pass: true)
    else
      create_test_suite_badge(pass: false)
    end
    measure_test_assertions()
  end
end


=begin
FUNCTIONS
=end

# @param _64 [boolean], @param coverage [boolean]
# The binary files with 0 bytes fill the purpose of showing if the exe was compiled with x86 or x64 compiler.
def build_unit_tests(_64: false, coverage: false)
  command = []
  if(_64)
    File.delete('temp/x64') rescue nil
    # locally using mingw compiler that comes packed with RubyInstaller 3.0 (version 10.2)
    command << gcc = 'C:/mingw64/bin/gcc.exe'
  else
    File.delete('temp/x86') rescue nil
    # locally using mingw compiler downloaded from official website (version 9.2)
    command << gcc = 'gcc'
  end
  command << debug = '-ggdb' # prepare exe for debugging purpose (gdb)
  if(coverage)
    command << coverage = '--coverage' # code coverage support (gcov)
  end
  command << standard = '-std=c11'
  command << warnings = '-w' # libraries use to throw some warnings
  command << preprocessing = '-D UNITY_INCLUDE_DOUBLE -D UNITY_SUPPORT_64' # double & long long support
  command << directories_included_for_headers_search = '-I ext/ -I src/ -I test/'
  command << source_files = 'ext/zip.c ext/sxmlc.c ext/sxmlsearch.c src/xlsx_drone.c ext/unity.c test/xlsx_drone.test.c'
  command << output = '-o temp/unit_tests.exe'
  # output command to trigger
  puts command.join(' ')
  # will return true if were no problems
  if(system(command.join(' ')))
    print("\nCompile successful")
    if(_64)
      File.open('temp/x64', 'wb') {|f|}
      print("(x64).\n\n")
    else
      File.open('temp/x86', 'wb') {|f|}
      print("(x86).\n\n")
    end
      true
  else
    puts("\nCompile failed.")
    false
  end
end

# @return [boolean]
def run_unit_tests
  result = system("temp/unit_tests")
  puts()
  result
end

def clear_unit_tests
  File.delete('temp/unit_tests.exe') rescue nil
  File.delete('temp/x86') rescue nil
  File.delete('temp/x64') rescue nil
  puts('Clear completed.')
end

# Executes gcov for the library source code according to the test suite. Prepares a JSON to be consumed by shields.io and then shown in README.md.
def measure_code_coverage
  # at this point *.gcda & *.gcno files gets generated in project root dir
  temp_gcov_path = 'temp/gcov.txt'
  system("gcov -o . src/xlsx_drone.c > #{temp_gcov_path}")
  # that generated several *.gcov files, what matters is the command output
  covered = \
    File.open(temp_gcov_path) do |f|
      f.read.match(/executed:(\d+[.,]?\d{0,2}%)/).[](1).sub(/,/, '.') # i.e.: "32.43%"
    end
  # create the *.json file rdy to be consumed by shields.io
  require 'json'
  # inquire matching color according to the amount of code covered
  color = \
    case (covered.to_f)
      when (80..100)
        'brightgreen'
      when (60..80)
        'green'
      when (40..60)
        'yellowgreen'
      else
        'orange'
    end
  # produce the json string
  json = {
    'schemaVersion': 1,
    'label': 'coverage',
    'message': covered,
    'color': color
  }.to_json
  # write the file
  File.open('data/shields/gcov.json', 'w:utf-8:utf-8') {|f| f.write(json)}
  # remove garbage
  require 'fileutils'
  FileUtils.remove(Dir['*.{gcda,gcno,gcov}'] << temp_gcov_path, force: true)
  # provide some output
  puts("Done. Code coverage: #{covered}.")
end

# Measure the amount of test assertions written and produce a JSON file rdy to be consumed by shields.io.
def measure_test_assertions
  # collect the data
  assertions = File.open('test/xlsx_drone.test.c') {|f| f.read.scan(/TEST_ASSERT/).size.to_s}
  # produce the json string
  require 'json'
  json = {
    'schemaVersion': 1,
    'label': 'test assertions',
    'message': assertions,
    'color': 'informational'
  }.to_json
  # write the file
  File.open('data/shields/assertions.json', 'w:utf-8:utf-8') {|f| f.write(json)}
  # provide some output
  puts("Assertions: #{assertions}.")
end

# @param pass [boolean]
def create_test_suite_badge(pass:)
  # produce the json string
  require 'json'
  json = {
    'schemaVersion': 1,
    'label': 'test suite',
    'message': pass ? 'pass' : 'fail',
    'color': pass ? 'brightgreen' : 'red'
  }.to_json
  # write the file
  File.open('data/shields/test_suite.json', 'w:utf-8:utf-8') {|f| f.write(json)}
end
