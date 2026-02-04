package document

// TrackChangesManager provides track change functionality.
type TrackChangesManager struct {
	doc *documentImpl
}

// Enabled reports whether track changes is enabled.
func (t *TrackChangesManager) Enabled() bool {
	if t == nil || t.doc == nil {
		return false
	}
	return t.doc.TrackChangesEnabled()
}

// Enable enables track changes.
func (t *TrackChangesManager) Enable() {
	if t == nil || t.doc == nil {
		return
	}
	t.doc.EnableTrackChanges(t.doc.trackAuthor)
}

// Disable disables track changes.
func (t *TrackChangesManager) Disable() {
	if t == nil || t.doc == nil {
		return
	}
	t.doc.DisableTrackChanges()
}

// Author returns the track changes author.
func (t *TrackChangesManager) Author() string {
	if t == nil || t.doc == nil {
		return ""
	}
	return t.doc.TrackAuthor()
}

// SetAuthor sets the track changes author.
func (t *TrackChangesManager) SetAuthor(name string) {
	if t == nil || t.doc == nil {
		return
	}
	t.doc.SetTrackAuthor(name)
}

// Insertions returns tracked insertions.
func (t *TrackChangesManager) Insertions() []Revision {
	if t == nil || t.doc == nil {
		return nil
	}
	revisions := t.doc.Insertions()
	result := make([]Revision, len(revisions))
	for i, rev := range revisions {
		result[i] = rev
	}
	return result
}

// Deletions returns tracked deletions.
func (t *TrackChangesManager) Deletions() []Revision {
	if t == nil || t.doc == nil {
		return nil
	}
	revisions := t.doc.Deletions()
	result := make([]Revision, len(revisions))
	for i, rev := range revisions {
		result[i] = rev
	}
	return result
}

// AllRevisions returns all tracked revisions.
func (t *TrackChangesManager) AllRevisions() []Revision {
	if t == nil || t.doc == nil {
		return nil
	}
	revisions := t.doc.AllRevisions()
	result := make([]Revision, len(revisions))
	for i, rev := range revisions {
		result[i] = rev
	}
	return result
}

// AcceptAll accepts all revisions.
func (t *TrackChangesManager) AcceptAll() {
	if t == nil || t.doc == nil {
		return
	}
	t.doc.AcceptAllRevisions()
}

// RejectAll rejects all revisions.
func (t *TrackChangesManager) RejectAll() {
	if t == nil || t.doc == nil {
		return
	}
	t.doc.RejectAllRevisions()
}

// AcceptRevision accepts a revision by ID.
func (t *TrackChangesManager) AcceptRevision(id string) error {
	if t == nil || t.doc == nil {
		return ErrInvalidIndex
	}
	for _, rev := range t.doc.AllRevisions() {
		if rev.ID() == id {
			return rev.Accept()
		}
	}
	return ErrInvalidIndex
}

// RejectRevision rejects a revision by ID.
func (t *TrackChangesManager) RejectRevision(id string) error {
	if t == nil || t.doc == nil {
		return ErrInvalidIndex
	}
	for _, rev := range t.doc.AllRevisions() {
		if rev.ID() == id {
			return rev.Reject()
		}
	}
	return ErrInvalidIndex
}

// InsertText inserts tracked text at a position.
func (t *TrackChangesManager) InsertText(para Paragraph, position int, text string) error {
	if t == nil || t.doc == nil || para == nil {
		return ErrInvalidIndex
	}
	if position < 0 {
		position = 0
	}
	runs := para.Runs()
	if position >= len(runs) {
		para.InsertTrackedText(text)
		return nil
	}
	para.InsertTrackedText(text)
	return nil
}

// DeleteText deletes tracked text by run range.
func (t *TrackChangesManager) DeleteText(para Paragraph, start, end int) error {
	if t == nil || t.doc == nil || para == nil {
		return ErrInvalidIndex
	}
	for i := end; i >= start; i-- {
		if err := para.DeleteTrackedText(i); err != nil {
			return err
		}
	}
	return nil
}

// ReplaceText replaces text in a paragraph.
func (t *TrackChangesManager) ReplaceText(para Paragraph, oldText, newText string) error {
	if t == nil || t.doc == nil || para == nil {
		return ErrInvalidIndex
	}
	if oldText == "" {
		return ErrInvalidIndex
	}
	for _, run := range para.Runs() {
		if run.Text() == oldText {
			run.SetText(newText)
			return nil
		}
	}
	return ErrInvalidIndex
}
