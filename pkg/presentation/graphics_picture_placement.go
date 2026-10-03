package presentation

import (
	"math"
	"math/big"
)

type PictureIntrinsicSize struct {
	Width  int64 `json:"width"`
	Height int64 `json:"height"`
}
type PicturePlacement struct {
	Geometry PictureGeometry `json:"geometry"`
	Crop     PictureCrop     `json:"crop"`
}
type FittedPictureOptions struct {
	PictureOptions
	Fit       string               `json:"fit"`
	Intrinsic PictureIntrinsicSize `json:"intrinsic"`
}
type FittedPictureReceipt struct {
	PictureReceipt
	PicturePlacement
}

// CalculatePicturePlacement uses exact rational intermediates and explicit rounding.
func CalculatePicturePlacement(box PictureGeometry, intrinsic PictureIntrinsicSize, fit string) (PicturePlacement, error) {
	var empty PicturePlacement
	if box.X < math.MinInt32 || box.X > math.MaxInt32 || box.Y < math.MinInt32 || box.Y > math.MaxInt32 || box.Width < 1 || box.Width > math.MaxInt32 || box.Height < 1 || box.Height > math.MaxInt32 || intrinsic.Width < 1 || intrinsic.Width > math.MaxInt32 || intrinsic.Height < 1 || intrinsic.Height > math.MaxInt32 {
		return empty, graphicsUnsupported("fit dimensions/coordinate bounds")
	}
	if fit != "contain" && fit != "cover" && fit != "stretch" {
		return empty, graphicsUnsupported("fit policy")
	}
	result := PicturePlacement{Geometry: box}
	wide := intrinsic.Width*box.Height > intrinsic.Height*box.Width
	if fit == "contain" {
		if wide {
			result.Geometry.Height = box.Width * intrinsic.Height / intrinsic.Width
		} else {
			result.Geometry.Width = box.Height * intrinsic.Width / intrinsic.Height
		}
		if result.Geometry.Width < 1 || result.Geometry.Height < 1 {
			return empty, graphicsUnsupported("contained extent rounds below one EMU")
		}
		result.Geometry.X += (box.Width - result.Geometry.Width) / 2
		result.Geometry.Y += (box.Height - result.Geometry.Height) / 2
		if result.Geometry.X > math.MaxInt32 || result.Geometry.Y > math.MaxInt32 {
			return empty, graphicsUnsupported("contained coordinate overflow")
		}
	}
	if fit == "cover" {
		denominator := intrinsic.Width * box.Height
		numerator := denominator - intrinsic.Height*box.Width
		if !wide {
			denominator = intrinsic.Height * box.Width
			numerator = denominator - intrinsic.Width*box.Height
		}
		n := new(big.Int).Mul(big.NewInt(numerator), big.NewInt(100000))
		n.Add(n, big.NewInt(denominator))
		d := new(big.Int).Mul(big.NewInt(denominator), big.NewInt(2))
		side := new(big.Int).Quo(n, d).Int64()
		if side*2 >= 100000 {
			return empty, graphicsUnsupported("cover rounds to empty source")
		}
		if wide {
			result.Crop.Left = side
			result.Crop.Right = side
		} else {
			result.Crop.Top = side
			result.Crop.Bottom = side
		}
	}
	return result, nil
}
func (s *EditSession) AddFittedPicture(part string, payload []byte, box PictureGeometry, options FittedPictureOptions) (FittedPictureReceipt, error) {
	placement, err := CalculatePicturePlacement(box, options.Intrinsic, options.Fit)
	if err != nil {
		return FittedPictureReceipt{}, err
	}
	receipt, err := s.addPicture(part, payload, placement.Geometry, options.PictureOptions, placement.Crop)
	if err != nil {
		return FittedPictureReceipt{}, err
	}
	return FittedPictureReceipt{receipt, placement}, nil
}
