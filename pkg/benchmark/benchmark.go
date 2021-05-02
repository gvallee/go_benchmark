//
// Copyright (c) 2021, NVIDIA CORPORATION. All rights reserved.
//
// See LICENSE.txt for license information
//

package benchmark

import "github.com/gvallee/go_software_build/pkg/app"

// Config represents the static OSU configuration (what never changes at runtime)
type Config struct {
	URL string

	Tarball string
}

// Install gathers all the data regarding the installation of OSU so it can easily be looked up later on
type Install struct {
	SubBenchmarks []app.Info
}
