################################################################################
# 
# NAME  : ps_grep.ps1
# AUTHOR: Seth Wright , James Madison University
# DATE  : 5/13/2009
# 
# Copyright (c) 2009 Seth Wright
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
################################################################################

if (Test-Path function:ps_grep) { Remove-Item function:ps_grep }

function global:ps_grep([switch] $c, [switch] $i, $pattern, $file="", $inputObject=$Null) {
  BEGIN {
    # This section executes before the pipeline.
    # $count will contain the number of matches found.
    $count = 0
    
    if ($inputObject) {
      $inputObject
    }
  } # end 'begin'

  PROCESS {
    # This section executes for each object in the pipeline.

    # Did we get a filename?  If so we're operating on a file, not stdout.
    if ($file -ne "") {

      # Is the filename a valid file?
      if (Test-Path $file) {

        # Grep the file for $pattern, if we were given $pattern.
        if($pattern) {
          if (!$i) {
            # User has specified as case-sensitive match.
            $case = "-CaseSensitive"
          }
          $command = "Select-String $case -pattern $pattern -InputObject (Get-Item $file)"
          foreach ($match in Invoke-Expression($command)) {
            if (!$c) {
              # only output lines if -c was not specified.
              Write-Host $match
            }
            # Increment our count.
            $count++
          }
        } # end 'if ($pattern)'
      } # end 'if (test-path)'

      else {
        # The filename passed does not exist.  Die.
        Write-Host "File does not exist."
      } # end 'else'
    } # end 'if ($file...)'

    # Didn't get a filename, so we're operating on stdout.
    elseif($pattern) {
      # Save the pipelined object into an internal var to avoid confusion below.
      $obj = $_

      # Loop through $obj's properties and look for $pattern
      foreach ($prop in ($obj | Get-Member -MemberType Properties) ) {

        if($obj.($prop.name.ToString()) -match $pattern) {

          if (!$c) {
            # Found a match, output the result to stdout.
            Write-Output $obj | &($MyInvocation.InvocationName) -inputObject $_
          }

          # Increment our count.
          $count++

          # No reason to keep searching this property.
          break
        } # end 'if ($obj...)'
      } # end 'foreach'
    } # end 'elseif'
  } # end 'process'

  END {
    # If '-c' was specified on the command line, output the total count.
    if ($c) {
      Write-Host $count
    } # end 'if ($c)'
  } # end 'end{}'
} # end function

Write-Host "Added ps_grep to global functions." -Fore White
