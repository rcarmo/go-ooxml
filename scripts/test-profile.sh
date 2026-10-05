#!/usr/bin/env bash
# Profile each Go package separately, retaining binaries/logs and post-run top tables.
set -uo pipefail
if (($# == 0)); then echo 'usage: scripts/test-profile.sh package [package ...]' >&2; exit 2; fi
repo=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd -P)
module=$(pwd -P)
case "$module" in "$repo"|"$repo/acceptance") ;; *) echo "run from Go module root or acceptance/: $module" >&2; exit 2;; esac
# Resolve before changing TMPDIR; standalone CI invocations use the same policy as Make.
source "$repo/scripts/project-tmp.sh"
project_tmp=$(project_tmp_resolve go-ooxml)
project_tmp_init "$project_tmp"
export PROJECT_TMP_ROOT="$project_tmp"
if [[ -n "${PROFILE_OUTPUT_DIR:-}" ]]; then
 case "$PROFILE_OUTPUT_DIR" in "$repo"/artifacts/*) ;; *) echo 'PROFILE_OUTPUT_DIR must be retained under repository artifacts/' >&2; exit 2;; esac
fi
if [[ -n "${PROFILE_JSON_EVENTS:-}" ]]; then
 case "$PROFILE_JSON_EVENTS" in "$repo"/artifacts/*) ;; *) echo 'PROFILE_JSON_EVENTS must be retained under repository artifacts/' >&2; exit 2;; esac
fi
run_id="$(date -u +%Y%m%dT%H%M%SZ)-$$"
module_name=root
[[ "$module" == "$repo/acceptance" ]] && module_name=acceptance
run="$project_tmp/runs/tests/$run_id/$module_name"
evidence="$repo/artifacts/profiles/$run_id/$module_name"
mkdir -p "$run" "$evidence" "$project_tmp/cache/go/build" "$project_tmp/cache/go/mod" "$project_tmp/cache/go/path" "$project_tmp/cache/xdg" "$project_tmp/cache/nuget" "$project_tmp/build/go"
export TMPDIR="$run" TMP="$run" TEMP="$run" GOTMPDIR="$project_tmp/build/go"
export GOCACHE="$project_tmp/cache/go/build" GOMODCACHE="$project_tmp/cache/go/mod" GOPATH="$project_tmp/cache/go/path" XDG_CACHE_HOME="$project_tmp/cache/xdg" NUGET_PACKAGES="$project_tmp/cache/nuget"
export GOMAXPROCS="${GOMAXPROCS:-2}"
package_list="$evidence/packages.txt"
if ! go list "$@" > "$package_list" 2> "$evidence/list.stderr"; then cat "$evidence/list.stderr" >&2; echo 'package discovery failed; no test profiles' >&2; exit 2; fi
mapfile -t packages < "$package_list"
if ((${#packages[@]} == 0)); then echo 'no Go packages selected; no test profiles' >&2; exit 2; fi
read -r -a extra_flags <<< "${PROFILE_GO_FLAGS:-}"
if [[ -n "${PROFILE_JSON_EVENTS:-}" ]]; then
 mkdir -p "$(dirname "$PROFILE_JSON_EVENTS")"
 : > "$PROFILE_JSON_EVENTS"
 extra_flags+=(-json)
fi
if [[ -n "${PROFILE_OUTPUT_DIR:-}" ]]; then
 mkdir -p "$PROFILE_OUTPUT_DIR"
fi
{
 printf 'revision=%s\nmodule=%s\ntoolchain=%s\nGOMAXPROCS=%s\nGOCACHE=%s\nGOMODCACHE=%s\nGOPATH=%s\nGOTMPDIR=%s\nTMPDIR=%s\n' "$(git -C "$repo" rev-parse HEAD)" "$module_name" "$(go version)" "$GOMAXPROCS" "$GOCACHE" "$GOMODCACHE" "$GOPATH" "$GOTMPDIR" "$TMPDIR"
 printf 'packages='; printf '%s ' "${packages[@]}"; printf '\nflags=-p 1 -count=1 -timeout 300s -cpuprofile -memprofile -outputdir\nextra_flags=%s\nCPU sampling=Go default (100 Hz)\nheap sampling=Go default (512 KiB average); inuse_space/alloc_space/alloc_objects\nworkload=uncached package tests; benchmark/fuzz/race and child-process activity only if explicitly requested and separately profiled\n' "${PROFILE_GO_FLAGS:-}"
} > "$evidence/command.txt"
status=0
index=0
for package in "${packages[@]}"; do
 index=$((index+1))
 path="$evidence/$(printf '%03d' "$index")-$(printf '%s' "$package" | tr '/.' '__')"
 mkdir -p "$path" "$run/pkg-$index"
 printf 'package=%s\n' "$package" > "$path/command.txt"
 printf 'go test -p 1 -count=1 -timeout 300s -cpuprofile %s -memprofile %s -outputdir %s' "$path/cpu.pprof" "$path/heap.pprof" "$path" >> "$path/command.txt"
 package_flags=("${extra_flags[@]}")
 if [[ "${PROFILE_COVERAGE:-0}" == 1 ]]; then package_flags+=(-coverprofile="$path/coverage.out"); fi
 printf ' %q' "${package_flags[@]}" "$package" >> "$path/command.txt"
 printf '\n' >> "$path/command.txt"
 export TMPDIR="$run/pkg-$index" TMP="$run/pkg-$index" TEMP="$run/pkg-$index"
 if [[ -n "${PROFILE_OUTPUT_DIR:-}" ]]; then export OOXML_GRAPHICS_OUTPUT="$PROFILE_OUTPUT_DIR"; fi
 binary="$path/package.test"
 # Build with matching instrumentation, not execution-only test/benchmark flags.
 compile_flags=()
 for flag in "${package_flags[@]}"; do
  case "$flag" in -race|-msan|-asan|-cover|-covermode=*|-coverpkg=*|-coverprofile=*) compile_flags+=("$flag");; esac
 done
 go test -c -o "$binary" "${compile_flags[@]}" "$package" > "$path/build.stdout" 2> "$path/build.stderr"
 build_rc=$?
 if [[ "$build_rc" != 0 ]]; then
  printf '%s\n' "$build_rc" > "$path/exit"
  cat "$path/build.stdout" "$path/build.stderr"
  echo "BUILD FAILED for $package (exit $build_rc); no CPU/heap profiles, binary or test execution" | tee "$path/profile-error.txt" >&2
  status=1
  continue
 fi
 go test -p 1 -count=1 -timeout 300s -cpuprofile "$path/cpu.pprof" -memprofile "$path/heap.pprof" -outputdir "$path" "${package_flags[@]}" "$package" > "$path/test.stdout" 2> "$path/test.stderr"
 rc=$?
 if [[ -n "${PROFILE_JSON_EVENTS:-}" ]]; then cat "$path/test.stdout" >> "$PROFILE_JSON_EVENTS"; fi
 printf '%s\n' "$rc" > "$path/exit"
 cat "$path/test.stdout" "$path/test.stderr"
 [[ "$rc" == 0 ]] || status=1
 if [[ "$rc" == 0 ]] && { grep -Fq '[no test files]' "$path/test.stdout" || { [[ "${PROFILE_COVERAGE:-0}" == 1 ]] && grep -Eq 'coverage: 0[.]0% of statements' "$path/test.stdout" && ! grep -q 'ok[[:space:]]' "$path/test.stdout"; }; }; then
  echo 'No test workload: compiled binary retained; CPU/heap profiles inapplicable' > "$path/no-tests.txt"
  continue
 fi
 if [[ ! -s "$binary" || ! -s "$path/cpu.pprof" || ! -s "$path/heap.pprof" ]]; then
  echo "PROFILE CAPTURE FAILED for $package (exit $rc); inspect retained logs/binary" | tee "$path/profile-error.txt" >&2
  status=1
  continue
 fi
 printf '%s\n' "$binary" > "$path/binary.txt"
 for kind in cpu alloc_space alloc_objects; do
  if [[ "$kind" == cpu ]]; then
   go tool pprof -cum -top -nodecount=25 "$binary" "$path/cpu.pprof" > "$path/$kind.cumulative.txt" 2> "$path/$kind.stderr" || status=1
  else
   go tool pprof -sample_index="$kind" -cum -top -nodecount=25 "$binary" "$path/heap.pprof" > "$path/$kind.cumulative.txt" 2> "$path/$kind.stderr" || status=1
  fi
 done
 if grep -Eq 'Total samples = 0([[:space:]]|$)|Total samples: 0([[:space:]]|$)|Total: 0([[:space:]]|$)' "$path/cpu.cumulative.txt"; then echo "EMPTY CPU SAMPLES for $package; representative workload required" | tee "$path/cpu-empty.txt" >&2; fi
 echo "PROFILE $package: $path (review cumulative CPU, alloc_space, alloc_objects; distinguish test/runtime from application hotspots)"
done
printf '%s\n' "$status" > "$evidence/exit"
if [[ -n "${PROFILE_JSON_EVENTS:-}" && "$status" != 0 ]]; then
 echo 'Go test/profile capture failed: graphics reconciliation must not consume this JSON' >&2
fi
echo "Retained profile evidence: $evidence; disposable scratch (not deleted): $run"
exit "$status"
