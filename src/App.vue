<script setup lang="ts">
import { ref } from 'vue'
import { addRow } from './excel'

const type = ref('NS&I')
const amount = ref(10)
const date = ref(new Date().toISOString())

const saving = ref(false)

function toExcelDate(date: Date): string {
  let converted =
    25569.0 +
    (date.getTime() - date.getTimezoneOffset() * 60 * 1000) /
      (1000 * 60 * 60 * 24)
  return converted.toString()
}

async function onSubmit(e: Event): Promise<void> {
  e.preventDefault()
  saving.value = true
  await addRow(
    toExcelDate(new Date(date.value)),
    type.value,
    amount.value.toString()
  )
  saving.value = false
}
</script>

<template>
  <div class="container px-3 py-4 mx-auto">
    <form v-on:submit="onSubmit">
      <div class="mb-3">
        <label for="type" class="form-label">Type</label>
        <input type="text" class="form-control" id="type" v-model="type" />
      </div>
      <div class="mb-3">
        <label for="amount" class="form-label">Amount (£)</label>
        <input
          type="number"
          class="form-control"
          id="amount"
          step="0.01"
          v-model="amount"
        />
      </div>

      <div class="mb-3">
        <a href="#collapseTarget" data-bs-toggle="collapse">Advanced</a>
      </div>
      <div class="collapse mb-3" id="collapseTarget">
        <label for="date" class="form-label">Date</label>
        <input type="text" class="form-control" id="date" v-model="date" />
      </div>

      <button type="submit" class="btn btn-primary" :disabled="saving">
        Submit
      </button>
    </form>
  </div>
</template>
